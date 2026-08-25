# -*- coding: utf-8 -*-
"""Generate harness pages that run the REAL extension code so screenshots show real UI."""
import pathlib, html

BASE = pathlib.Path(__file__).parent
APP = BASE / "app"

# ──────────────────────────────────────────────────────────────────────────────
# Demo documents
# ──────────────────────────────────────────────────────────────────────────────

DOC_API = r"""
# Payments API — Design Notes

> Rendered **instantly** from a plain `.md` file on your own disk. No account, no upload, no telemetry.

## Overview

The payments service exposes a small, versioned HTTP surface. Every request is
idempotent, every response is JSON, and every mutation is written to the ledger
*before* it is acknowledged to the caller.

## Quick start

```bash
npm install @acme/payments
export ACME_API_KEY="sk_live_9f2c...804a"
```

```typescript
import { Payments } from '@acme/payments';

const client = new Payments({ apiKey: process.env.ACME_API_KEY });

export async function charge(customer: string, amount: number) {
  const intent = await client.intents.create({
    customer,
    amount,                 // minor units — 4999 = $49.99
    currency: 'usd',
    capture: true,
  });

  if (intent.status !== 'succeeded') {
    throw new Error(`Charge failed: ${intent.lastError}`);
  }
  return intent.id;
}
```

## Endpoints

| Method | Path | Description | Auth |
| :--- | :--- | :--- | :---: |
| `POST` | `/v2/intents` | Create a payment intent | Yes |
| `GET` | `/v2/intents/:id` | Fetch a single intent | Yes |
| `POST` | `/v2/refunds` | Refund a captured charge | Yes |
| `GET` | `/v2/balance` | Current available balance | Yes |

## Rollout checklist

- [x] Double-entry ledger migration
- [x] Idempotency keys on every mutation
- [ ] Regional failover for `eu-west-1`
- [ ] Public changelog and SDK release notes

## Error handling

Errors carry a stable machine-readable `code`. Retry only on `rate_limited`
and `upstream_timeout`; everything else is terminal.

### Retry policy

Exponential backoff starting at 250 ms, capped at 8 s, with full jitter.
"""

DOC_RELEASE = r"""
# Release 2.4 — Notes

A maintenance release focused on the editor, plus two long-requested exports.

## Highlights

- **Split view** now keeps both panes scroll-locked while you type
- **Format toolbar** gained table and task-list buttons
- Documents export to Word with diagrams and math intact

## Editor

The editor is a plain textarea with smart behaviour layered on top: `Tab` and
`Shift+Tab` indent selections, `Enter` continues the list you are in, and
`Ctrl+S` writes the file straight back to disk.

| Shortcut | Action |
| :--- | :--- |
| `Ctrl+B` | Bold |
| `Ctrl+I` | Italic |
| `Ctrl+K` | Insert link |
| `Ctrl+S` | Save file |

## Upgrade notes

- [x] No configuration changes required
- [ ] Re-pin the extension if your toolbar was customised

```js
// Preview stays in sync as you type
editor.addEventListener('input', () => {
  preview.innerHTML = parseMarkdown(editor.value);
});
```
"""

DOC_DIAGRAM = r"""
# Request Lifecycle

```mermaid
graph LR
    A[Browser] --> B{Edge cache}
    B -- hit --> C[Return cached]
    B -- miss --> D[Payments API]
    D --> E[(Ledger)]
```

## Latency budget

The p95 wall-clock budget is 400 ms. Across $n$ shards the expected fan-out cost is:

$$
T_{p95} \;=\; T_{fixed} \;+\; \max_{i \le n}\bigl(T_i\bigr) \;+\; \frac{\sigma}{\sqrt{n}}
$$

## Sequence

```mermaid
sequenceDiagram
    participant C as Client
    participant G as Gateway
    participant S as Service
    C->>G: POST /v2/intents
    G->>S: create(intent)
    S-->>G: 201 Created
    G-->>C: intent.id
```

## Notes

Mermaid and KaTeX both ship inside the extension, so diagrams and formulas
render offline — nothing is fetched from a CDN.
"""

DOC_SEARCH = r"""
# Webhook Guide

Webhook delivery is at-least-once. Your endpoint must be idempotent, because a
webhook can arrive twice if our first delivery attempt times out.

## Registering a webhook

Register one webhook per environment. Each webhook gets its own signing secret,
and rotating a secret never drops in-flight deliveries.

```bash
curl -X POST https://api.acme.dev/v2/webhooks \
  -d url="https://example.com/hooks/acme" \
  -d events="intent.succeeded,refund.created"
```

## Verifying a webhook

Every webhook request carries an `Acme-Signature` header. Compute the HMAC of
the raw body and compare in constant time.

| Header | Meaning |
| :--- | :--- |
| `Acme-Signature` | HMAC-SHA256 of the raw webhook body |
| `Acme-Timestamp` | Unix seconds — reject anything older than 5 minutes |
| `Acme-Delivery` | Unique id for this webhook delivery attempt |

## Retries

A webhook that does not return `2xx` within 10 seconds is retried with
exponential backoff for up to 24 hours.

- [x] Respond `200` before doing slow work
- [ ] Alert when the webhook queue depth exceeds 1000
"""

# ──────────────────────────────────────────────────────────────────────────────

SHIM = """
window.__MD_HARNESS__ = true;
window.__MD_FILENAME__ = %(fname)s;
window.__MD_THEME__ = %(theme)s;
window.chrome = {
  runtime: { id: 'harness', getURL: function (p) { return p; } },
  storage: { local: {
    get: function (k, cb) { cb({ theme: window.__MD_THEME__, fontSize: 16 }); },
    set: function () {}
  } },
  tabs: { query: function (q, cb) { cb([{ id: 1, url: 'file:///C:/docs/README.md' }]); },
          create: function () {} },
  scripting: { executeScript: function () { return Promise.resolve(); } }
};
"""

PAGE = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>%(title)s</title>
<style>html,body{margin:0;padding:0;background:%(bg)s}body>pre{visibility:hidden}</style>
</head>
<body>
<pre>%(md)s</pre>
<script>%(shim)s</script>
<script src="katex/katex.min.js"></script>
<script src="docx-export.js"></script>
<script src="content.mv.js"></script>
<script>
document.addEventListener('DOMContentLoaded', function () {
  var l = document.createElement('link');
  l.rel = 'stylesheet';
  l.href = 'content.css';
  document.head.appendChild(l);
});
window.addEventListener('load', function () {
  setTimeout(function () {
    try { %(actions)s } catch (e) { console.error('action failed', e); }
    document.title = 'READY';
  }, 120);
});
</script>
</body>
</html>
"""

SHOTS = [
    dict(out="app-view-dark.html",   md=DOC_API,     fname="payments-api.md",
         theme="dark",  bg="#0d1117", actions=""),
    dict(out="app-split-light.html", md=DOC_RELEASE, fname="release-2.4.md",
         theme="light", bg="#ffffff",
         actions="document.querySelector('.mode-btn[data-mode=\\\"split\\\"]').click();"),
    dict(out="app-diagram-dark.html", md=DOC_DIAGRAM, fname="architecture.md",
         theme="dark", bg="#0d1117", actions=""),
    dict(out="app-search-light.html", md=DOC_SEARCH, fname="webhook-guide.md",
         theme="light", bg="#ffffff",
         actions=("document.getElementById('md-content').scrollTop=516;"
                  "document.getElementById('btn-search').click();"
                  "var i=document.getElementById('search-input');"
                  "i.value='webhook';"
                  "i.dispatchEvent(new Event('input',{bubbles:true}));"
                  "i.blur();")),
    dict(out="app-theme-dark.html", md=DOC_API, fname="payments-api.md",
         theme="dark", bg="#0d1117",
         actions="document.getElementById('md-content').scrollTop=1062;"),
]

for s in SHOTS:
    body = PAGE % dict(
        title=s["fname"],
        bg=s["bg"],
        md=html.escape(s["md"].strip("\n"), quote=False),
        shim=SHIM % dict(fname='"%s"' % s["fname"], theme='"%s"' % s["theme"]),
        actions=s["actions"],
    )
    (APP / s["out"]).write_text(body, encoding="utf-8")
    print("wrote", s["out"])

# ── popup harness (dark) ──────────────────────────────────────────────────────
popup_src = (APP / "popup.html").read_text(encoding="utf-8")
DARK_VARS = """
<style>
:root{
  --bg:#0d1117; --bg-alt:#161b22; --border:#30363d; --text:#e6edf3;
  --muted:#7d8590; --link:#58a6ff; --accent:#58a6ff;
  --pill-active-bg:#388bfd; --pill-active-text:#fff;
  --pill-bg:#21262d; --pill-border:#30363d; --dot-ok:#3fb950; --dot-off:#6e7681;
}
html,body{background:var(--bg)}
</style>
<script>
window.chrome = {
  runtime: { id: 'harness' },
  storage: { local: {
    get: function (k, cb) { cb({ theme: 'dark', fontSize: 16 }); },
    set: function () {}
  } },
  tabs: { query: function (q, cb) { cb([{ id: 1, url: 'file:///C:/docs/README.md' }]); },
          create: function () {} },
  scripting: { executeScript: function () { return Promise.resolve(); } }
};
</script>
"""
popup_dark = popup_src.replace("</head>", DARK_VARS + "</head>")
(APP / "popup-dark.html").write_text(popup_dark, encoding="utf-8")
print("wrote popup-dark.html")
