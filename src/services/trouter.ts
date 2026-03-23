/**
 * Trouter WebSocket client for real-time Teams message notifications.
 * Based on the purple-teams reverse-engineered protocol.
 *
 * Flow:
 *   1. Get Skype token (OAuth with Teams desktop client ID → Teams authsvc)
 *   2. POST Trouter info → get socketio URL, surl, connectparams
 *   3. Socket.io v1 handshake → session ID
 *   4. WebSocket connect → send user.authenticate → register endpoints
 *   5. Receive 3:: push frames with NewMessage events
 */

import { EventEmitter } from "node:events";
import { promises as fs } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import { type AccountInfo, type Configuration, PublicClientApplication } from "@azure/msal-node";
import WebSocket from "ws";

// Teams desktop native client (public, has api.spaces.skype.com scope)
const TEAMS_CLIENT_ID = "1fec8e78-bce4-4aaf-ab1b-5451cc387264";
const SKYPE_SCOPES = ["https://api.spaces.skype.com/.default", "openid", "offline_access"];
const TROUTER_INFO_URL = "https://go.trouter.teams.microsoft.com/v4/a";
const REGISTRAR_URL = "https://teams.microsoft.com/registrar/prod/V2/registrations";
const TC_JSON = '{"cv":"2024.23.01.2","ua":"TeamsCDL","hr":"","v":"27/1.0.0.2023052414"}';

const SKYPE_MSAL_CACHE_PATH = join(homedir(), ".teams-mcp-skype-cache.json");

export interface TeamsMessage {
  id: string;
  from: string;
  body: string;
  chat_jid: string;
  chat_name: string;
  timestamp: number;
}

interface TrouterInfo {
  socketio: string;
  surl: string;
  connectparams: Record<string, string>;
  ccid: string;
}

export class TrouterClient extends EventEmitter {
  private accessToken = "";
  private skypeToken = "";
  private epid = crypto.randomUUID();
  private ws: WebSocket | null = null;
  private pingCount = 0;
  private pingInterval: ReturnType<typeof setInterval> | null = null;
  private reconnectTimeout: ReturnType<typeof setTimeout> | null = null;
  private running = false;
  private tenantId: string;
  private msalClient: PublicClientApplication | null = null;
  private cachedAccount: AccountInfo | null = null;
  // Maps Trouter-specific chat IDs to Graph chat IDs (e.g. "48:notes" → "19:xxx@unq.gbl.spaces")
  private chatIdAliases: Record<string, string> = {};

  constructor(tenantId: string) {
    super();
    this.tenantId = tenantId;
  }

  /** Register a chat ID alias. Trouter uses different IDs than Graph for some chats
   *  (e.g. self-chat is "48:notes" in Trouter but "19:xxx@unq.gbl.spaces" in Graph). */
  setChatIdAlias(trouterId: string, graphId: string): void {
    this.chatIdAliases[trouterId] = graphId;
  }

  /** Start the Trouter client. Connects and auto-reconnects. */
  async start(): Promise<void> {
    this.running = true;
    await this.initMsal();
    this.connectLoop();
  }

  stop(): void {
    this.running = false;
    if (this.pingInterval) clearInterval(this.pingInterval);
    if (this.reconnectTimeout) clearTimeout(this.reconnectTimeout);
    if (this.ws) this.ws.close();
  }

  /** Check if we have a cached MSAL account (i.e. user has logged in). */
  async hasAuth(): Promise<boolean> {
    try {
      await this.initMsal();
      if (!this.msalClient) return false;
      const accounts = await this.msalClient.getTokenCache().getAllAccounts();
      return accounts.length > 0;
    } catch {
      return false;
    }
  }

  /** Interactive device code login for Skype scope. Must be called once. */
  async login(): Promise<string> {
    await this.initMsal();
    if (!this.msalClient) throw new Error("MSAL client not initialized");

    return new Promise((resolve, reject) => {
      this.msalClient!.acquireTokenByDeviceCode({
        scopes: SKYPE_SCOPES,
        deviceCodeCallback: (resp) => {
          resolve(resp.message);
        },
      })
        .then(async (result) => {
          if (!result?.accessToken) {
            reject(new Error("No access token received"));
            return;
          }
          this.accessToken = result.accessToken;
          this.cachedAccount = result.account;
          this.emit("authenticated");
        })
        .catch(reject);
    });
  }

  /** Remove the Skype MSAL cache file. Called from logout(). */
  static async clearCache(): Promise<void> {
    try {
      await fs.unlink(SKYPE_MSAL_CACHE_PATH);
    } catch {
      // Ignore if file doesn't exist
    }
  }

  // --- Private ---

  private async initMsal(): Promise<void> {
    if (this.msalClient) return;

    const msalConfig: Configuration = {
      auth: {
        clientId: TEAMS_CLIENT_ID,
        authority: `https://login.microsoftonline.com/${this.tenantId}`,
      },
      cache: {
        cachePlugin: {
          async beforeCacheAccess(ctx) {
            try {
              const data = await fs.readFile(SKYPE_MSAL_CACHE_PATH, "utf8");
              ctx.tokenCache.deserialize(data);
            } catch (error) {
              if ((error as NodeJS.ErrnoException).code !== "ENOENT") {
                console.error("Warning: Could not read Skype token cache:", error);
              }
            }
          },
          async afterCacheAccess(ctx) {
            if (ctx.cacheHasChanged) {
              try {
                const data = ctx.tokenCache.serialize();
                await fs.writeFile(SKYPE_MSAL_CACHE_PATH, data, "utf8");
              } catch (error) {
                console.error("Warning: Could not write Skype token cache:", error);
              }
            }
          },
        },
      },
    };
    this.msalClient = new PublicClientApplication(msalConfig);
  }

  private async refreshTokens(): Promise<void> {
    if (!this.msalClient) await this.initMsal();

    // Get cached account
    if (!this.cachedAccount) {
      const accounts = await this.msalClient!.getTokenCache().getAllAccounts();
      if (accounts.length === 0) {
        throw new Error("No cached account. Run 'trouter-login' first.");
      }
      this.cachedAccount = accounts[0];
    }

    // Use acquireTokenSilent to refresh via MSAL cache
    const result = await this.msalClient!.acquireTokenSilent({
      scopes: SKYPE_SCOPES,
      account: this.cachedAccount,
    });

    if (!result?.accessToken) {
      throw new Error("Token refresh failed — no access token. Re-run 'trouter-login'.");
    }

    this.accessToken = result.accessToken;
    this.cachedAccount = result.account;
  }

  private async getSkypeToken(): Promise<string> {
    if (!this.accessToken) await this.refreshTokens();

    const resp = await fetch("https://teams.microsoft.com/api/authsvc/v1.0/authz", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${this.accessToken}`,
        "Content-Type": "application/json",
      },
      body: "",
    });

    const data = (await resp.json()) as Record<string, unknown>;
    const tokens = data.tokens as Record<string, unknown> | undefined;
    if (!tokens?.skypeToken) {
      throw new Error(`No skype token: ${JSON.stringify(data).slice(0, 200)}`);
    }

    this.skypeToken = tokens.skypeToken as string;
    return this.skypeToken;
  }

  private async connectLoop(): Promise<void> {
    while (this.running) {
      try {
        await this.connect();
      } catch (err) {
        console.error(`Trouter connect error: ${err}, retrying in 10s`);
      }
      if (this.running) {
        await new Promise((r) => {
          this.reconnectTimeout = setTimeout(r, 10000);
        });
      }
    }
  }

  private async connect(): Promise<void> {
    await this.refreshTokens();
    const skypeToken = await this.getSkypeToken();

    // 1. Trouter info
    const infoResp = await fetch(`${TROUTER_INFO_URL}?epid=${this.epid}`, {
      method: "POST",
      headers: { "X-Skypetoken": skypeToken, "Content-Length": "0" },
    });
    const info = (await infoResp.json()) as TrouterInfo;
    if (!info.socketio) info.socketio = "https://go.trouter.teams.microsoft.com/";

    console.error(`Trouter: socketio=${info.socketio} ccid=${info.ccid}`);

    // 2. Socket.io handshake
    const params = new URLSearchParams({
      v: "v4",
      tc: TC_JSON,
      con_num: `${this.epid}_1`,
      epid: this.epid,
      ccid: info.ccid,
      auth: "true",
      timeout: "40",
      ...info.connectparams,
    });

    const hsResp = await fetch(`${info.socketio}socket.io/1/?${params}`, {
      headers: { "X-Skypetoken": skypeToken },
    });
    const hsText = await hsResp.text();
    const sessionId = hsText.split(":")[0];
    console.error(`Trouter: session=${sessionId}`);

    // 3. WebSocket connect
    const wsScheme = info.socketio.startsWith("https") ? "wss" : "ws";
    const host = new URL(info.socketio).host;
    const wsUrl = `${wsScheme}://${host}/socket.io/1/websocket/${sessionId}?${params}`;

    return new Promise<void>((resolve, reject) => {
      const ws = new WebSocket(wsUrl, {
        headers: {
          "X-Skypetoken": skypeToken,
          "User-Agent": "Mozilla/5.0 Teams",
        },
      });
      this.ws = ws;

      ws.on("open", () => console.error("Trouter: WebSocket connected"));

      ws.on("message", (raw) => {
        const frame = raw.toString();
        this.handleFrame(ws, frame, skypeToken, info);
      });

      ws.on("close", () => {
        this.stopPing();
        resolve();
      });

      ws.on("error", (err) => {
        this.stopPing();
        reject(err);
      });

      // Start keepalive
      this.startPing(ws);

      // Register endpoints
      this.register(info, skypeToken);
    });
  }

  private handleFrame(ws: WebSocket, frame: string, _skypeToken: string, info: TrouterInfo): void {
    if (frame === "1::" || frame === "1:::") {
      console.error("Trouter: connected, authenticating");
      this.sendAuthenticate(ws, info);
      return;
    }

    if (frame === "2::" || frame === "2:::") {
      ws.send("2::");
      return;
    }

    if (frame.startsWith("3::")) {
      const payload = frame.startsWith("3:::") ? frame.slice(4) : frame.slice(3);
      this.handleNotification(ws, payload);
      return;
    }

    if (frame.startsWith("5:")) {
      const idx = frame.indexOf("::{");
      if (idx >= 0) {
        try {
          const evt = JSON.parse(frame.slice(idx + 2)) as { name: string; args?: unknown[] };
          if (evt.name === "trouter.connected") {
            console.error("Trouter: connected event");
          } else if (evt.name === "trouter.message_loss") {
            // Normal after reconnect, ignore
          }
        } catch {
          /* ignore */
        }
      }
    }
  }

  private sendAuthenticate(ws: WebSocket, info: TrouterInfo): void {
    const payload = {
      name: "user.authenticate",
      args: [
        {
          headers: {
            "X-Ms-Test-User": "False",
            Authorization: `Bearer ${this.accessToken}`,
            "X-MS-Migration": "True",
          },
          connectparams: info.connectparams,
        },
      ],
    };
    ws.send(`5:::${JSON.stringify(payload)}`);
  }

  private handleNotification(ws: WebSocket, payload: string): void {
    try {
      const frame = JSON.parse(payload) as {
        id: number;
        body?: string;
      };

      // ACK
      ws.send(`3:::${JSON.stringify({ id: frame.id, status: 200, body: "" })}`);

      if (!frame.body) return;

      const event = JSON.parse(frame.body) as {
        type?: string;
        resourceType?: string;
        resource?: Record<string, unknown>;
      };

      if (event.resourceType !== "NewMessage") return;

      const res = event.resource;
      if (!res) return;

      // Extract chat ID from conversationLink
      let chatId = (res.conversationLink as string) ?? "";
      const lastSlash = chatId.lastIndexOf("/");
      if (lastSlash >= 0) chatId = chatId.slice(lastSlash + 1);

      // Extract sender from "from" URL
      let from = (res.from as string) ?? "";
      const fromSlash = from.lastIndexOf("/");
      if (fromSlash >= 0) from = from.slice(fromSlash + 1);

      // Strip HTML from content
      let body = (res.content as string) ?? "";
      body = body
        .replace(/<[^>]+>/g, "")
        .replace(/&amp;/g, "&")
        .replace(/&lt;/g, "<")
        .replace(/&gt;/g, ">")
        .replace(/&nbsp;/g, " ")
        .trim();

      const composetime = (res.composetime as string) ?? "";
      const ts = composetime ? new Date(composetime).getTime() / 1000 : Date.now() / 1000;

      // Resolve Trouter chat ID to Graph ID if alias exists.
      const resolvedChatId = this.chatIdAliases[chatId] ?? chatId;

      const msg: TeamsMessage = {
        id: (res.id as string) ?? (res.clientmessageid as string) ?? "",
        from,
        body,
        chat_jid: resolvedChatId,
        chat_name: (res.threadtopic as string) ?? "",
        timestamp: Math.floor(ts),
      };

      console.error(
        `Trouter: NEW MESSAGE from=${msg.from} chat=${msg.chat_jid} body="${msg.body.slice(0, 50)}"`
      );
      this.emit("message", msg);
    } catch (err) {
      console.error(`Trouter: parse notification error: ${err}`);
    }
  }

  private async register(info: TrouterInfo, skypeToken: string): Promise<void> {
    const registrations = [
      {
        appId: "NextGenCalling",
        templateKey: "DesktopNgc_2.3:SkypeNgc",
        suffix: "NGCallManagerWin",
      },
      { appId: "SkypeSpacesWeb", templateKey: "SkypeSpacesWeb_2.3", suffix: "SkypeSpacesWeb" },
      { appId: "TeamsCDLWebWorker", templateKey: "TeamsCDLWebWorker_2.1", suffix: "" },
    ];

    for (const reg of registrations) {
      const path = reg.suffix ? info.surl.replace(/\/$/, "") + "/" + reg.suffix : info.surl;

      const body = {
        clientDescription: {
          appId: reg.appId,
          aesKey: "",
          languageId: "en-US",
          platform: "edge",
          templateKey: reg.templateKey,
          platformUIVersion: "27/1.0.0.2023052414",
        },
        registrationId: this.epid,
        nodeId: "",
        transports: {
          TROUTER: [{ context: "", path, ttl: 86400 }],
        },
      };

      try {
        const resp = await fetch(REGISTRAR_URL, {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            "X-Skypetoken": skypeToken,
            Authorization: `Bearer ${this.accessToken}`,
          },
          body: JSON.stringify(body),
        });
        console.error(`Trouter: registered ${reg.appId} → ${resp.status}`);
      } catch (err) {
        console.error(`Trouter: register ${reg.appId} error: ${err}`);
      }
    }
  }

  private startPing(ws: WebSocket): void {
    this.pingCount = 0;
    this.pingInterval = setInterval(() => {
      this.pingCount++;
      ws.send(`5:${this.pingCount}+::${JSON.stringify({ name: "ping" })}`);
    }, 30000);
  }

  private stopPing(): void {
    if (this.pingInterval) {
      clearInterval(this.pingInterval);
      this.pingInterval = null;
    }
  }
}
