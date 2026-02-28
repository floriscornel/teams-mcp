# Copilot Instructions — teams-mcp

## Project Overview

MCP (Model Context Protocol) server that bridges AI assistants to Microsoft Teams via the Microsoft Graph API. Published as `@floriscornel/teams-mcp` on npm. Uses MSAL device-code auth with a file-based token cache (`~/.teams-mcp-token-cache.json`).

## Architecture

- **`src/index.ts`** — Dual-mode entrypoint: CLI commands (`authenticate`, `check`, `logout`) and MCP server (`startMcpServer`). Routes via `process.argv`. Supports `--read-only` flag and `TEAMS_MCP_READ_ONLY` env var.
- **`src/services/graph.ts`** — Singleton `GraphService` wrapping the Microsoft Graph client. Handles token acquisition (env `AUTH_TOKEN` or MSAL silent refresh). Exports `READ_ONLY_SCOPES` and `FULL_SCOPES`. Has a `readOnlyMode` property that controls which scopes are requested. All Graph API calls go through `graphService.getClient()`.
- **`src/tools/*.ts`** — Each file exports a `register*Tools(server, graphService, readOnly)` function that registers MCP tools via `server.tool(name, description, zodSchema, handler)`. Tools are grouped by domain: `auth`, `users`, `teams`, `chats`, `search`. When `readOnly` is true, write tools are skipped.
- **`src/types/graph.ts`** — Re-exports `@microsoft/microsoft-graph-types` and defines project-specific interfaces (`GraphApiResponse<T>`, `*Summary` types). All response shapes use optional properties with `| undefined` to handle Graph API variability.
- **`src/utils/`** — Stateless helpers: `markdown.ts` (Markdown→sanitized HTML via `marked`+`DOMPurify`), `html-to-markdown.ts` (Teams HTML→Markdown via `turndown` with custom rules for `<at>`, `<attachment>`, `<systemEventMessage>`), `attachments.ts` (image hosted content), `file-upload.ts` (file uploads with automatic simple ≤4MB / resumable >4MB split), `users.ts` (mention processing).
- **`src/msal-cache.ts`** — File-based `ICachePlugin` for MSAL token persistence.

## Commands

```bash
npm run build          # clean + tsc
npm run dev            # tsx watch src/index.ts
npm test               # vitest run
npm run test:watch     # vitest (watch mode)
npm run test:coverage  # vitest with v8 coverage (80% threshold on branches/functions/lines/statements)
npm run lint           # biome check src/
npm run lint:fix       # biome check --write --unsafe src/
npm run format         # biome format --write src/
npm run ci             # biome ci src/ (CI mode)
```

## CI/CD

- **CI** (`.github/workflows/ci.yml`): Runs on PRs and pushes to `main`. Matrix tests Node 20/22/24. Steps: `npm run ci` (Biome), `tsc`, `vitest --coverage` with JUnit XML. Uploads coverage to Codecov.
- **Release** (`.github/workflows/release.yml`): Triggered by `v*` tags. Runs lint + tests then `npm publish` with provenance. Bump version in `package.json` and `src/index.ts` (`McpServer` version string) before tagging.

## Code Style & Conventions

- **Formatter/Linter**: Biome (not ESLint/Prettier). 2-space indent, double quotes, semicolons, LF line endings, 100-char line width, ES5 trailing commas. Run `npm run lint:fix` to auto-fix.
- **Module system**: ESM (`"type": "module"`). Use `.js` extensions in imports even for `.ts` files (e.g., `import { GraphService } from "../services/graph.js"`).
- **TypeScript**: Strict mode with `exactOptionalPropertyTypes`, `noImplicitAny`, `noUnusedLocals`, `noUnusedParameters`. Target ES2022. Use `node:` protocol for Node.js built-ins (e.g., `import { join } from "node:path"`).
- **Schema validation**: Use `zod` for MCP tool parameter schemas, passed directly to `server.tool()`.
- **Error handling in tools**: Catch errors inside each tool handler, return `{ content: [{ type: "text", text: "❌ Error: ..." }] }` — never throw from a tool handler.
- **Graph API response typing**: Always cast Graph responses with `as GraphApiResponse<T>` or specific types. Check `response?.value?.length` before mapping.

## Testing

- **Framework**: Vitest with globals enabled (`describe`, `it`, `expect` available without import, though they're typically imported explicitly).
- **HTTP mocking**: MSW (`msw/node`) for Graph API calls. The global setup in `src/test-utils/vitest.setup.ts` starts/resets/closes the MSW server. Fixtures and default handlers live in `src/test-utils/setup.ts`.
- **Unit test mocking**: `vi.mock()` for modules, `vi.fn()` for functions. Tests mock `GraphService.getClient()` to return a mock `Client` with chainable `.api().get()`/`.post()` methods.
- **Test location**: Co-located `__tests__/` directories mirroring `src/` structure (e.g., `src/tools/__tests__/chats.test.ts` tests `src/tools/chats.ts`).
- **Coverage**: v8 provider, 80% global threshold. `index.ts` and `test-utils/` are excluded from coverage.
- **Auth in tests**: `GraphService` is a singleton — tests mock it via `vi.mock()` on `@azure/msal-node` and `@microsoft/microsoft-graph-client`, setting up `PublicClientApplication` / `Client.initWithMiddleware` stubs *before* importing the module under test. See `src/services/__tests__/graph.test.ts` for the pattern.

## Microsoft Graph Documentation (microsoftdocs/mcp)

The workspace has a `microsoftdocs/mcp` server configured (`.vscode/mcp.json`) that provides access to official Microsoft Learn documentation. Use it to look up Graph API endpoints, permissions, and SDK usage.

**Workflow:**
1. Use `microsoft_docs_search` to find relevant docs (returns up to 10 content chunks with title, URL, and excerpt).
2. Use `microsoft_code_sample_search` to find code examples (returns up to 20 samples; optional `language` filter).
3. Use `microsoft_docs_fetch` to retrieve the full content of a specific doc page when search results are insufficient.

**When to use:**
- Verifying Graph API endpoint paths, query parameters, or response shapes.
- Checking required permissions/scopes for a Graph API call.
- Looking up SDK patterns for `@microsoft/microsoft-graph-client`.
- Resolving questions about Microsoft Teams-specific Graph API behavior (e.g., message types, mention formats, hosted content).

## Adding a New MCP Tool

1. Add or extend a `register*Tools` function in `src/tools/`.
2. Define parameters with `zod` schemas and pass to `server.tool()`.
3. Use `await graphService.getClient()` then chain `.api(endpoint)` calls.
4. Return `{ content: [{ type: "text", text: JSON.stringify(result, null, 2) }] }`.
5. Add types to `src/types/graph.ts` if new Graph API shapes are needed.
6. Write tests in `src/tools/__tests__/` following the mock-client pattern.
7. Register the new tool group in `src/index.ts` → `startMcpServer()` if it's a new file.
