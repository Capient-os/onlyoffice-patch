# ONLYOFFICE table layout patch

Capient converts Word and PowerPoint files with the converter from ONLYOFFICE Docs
Community Edition (AGPL-3.0). This repository is the source of the one change we make
to it, published under AGPL-3.0 section 13.

`sdkjs-patch.ts` changes two places in ONLYOFFICE's table layout so that a table breaks
over pages as Microsoft Word breaks it:

1. Rows that "keep with next" carries to the next page get the repeated header row.
2. A table never leaves only its header rows at the foot of its first page; it moves to
   the next page whole, with a heading kept to it.

The readable change, against `word/Editor/Table/TableRecalculate.js` at
[ONLYOFFICE/sdkjs tag v10.0.0.169](https://github.com/ONLYOFFICE/sdkjs/blob/v10.0.0.169/word/Editor/Table/TableRecalculate.js)
(`CTable.prototype.private_RecalculatePage`), is in the file's header. The script applies
it to the minified `sdkjs/word/sdk-all.js` shipped in
`onlyoffice/documentserver:10.0.0@sha256:ca17889f10b11141cc6441836aebdd2835cf3e1b49681b2fd6e155c94b7850b3`
(build 169), and stops if the file does not match:

```bash
bun sdkjs-patch.ts /var/www/onlyoffice/documentserver/sdkjs/word/sdk-all.js
```

The converter starts from a V8 snapshot when one is present, so remove
`sdkjs/word/sdk-all.bin` and rebuild it with `x2t -create-js-snapshots` afterwards.

ONLYOFFICE is a registered trademark of Ascensio System SIA. This patch is not endorsed by
Ascensio System SIA.
