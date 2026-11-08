// Two changes to ONLYOFFICE's table layout (sdkjs, AGPL-3.0), so a table breaks over
// pages as Word breaks it. Run once at image build on the shipped, minified
// `sdkjs/word/sdk-all.js` of onlyoffice/documentserver 10.0.0 (build 169, pinned by
// digest in Dockerfile.vercel). Each change must match exactly once, or the build stops.
//
//   bun sdkjs-patch.ts <path to sdk-all.js>
//
// This file is the source of the modification. In the readable source,
// word/Editor/Table/TableRecalculate.js at tag v10.0.0.169, CTable.prototype.private_RecalculatePage:
//
// 1. Rows carried to the next page by "keep with next" get the repeated header row.
//    The forced-keep break (`if (nKeepNextForceRow > FirstRow && CurRow === nKeepNextForceRow)`)
//    gains the line the overflow break already has:
//        if (-1 === this.HeaderInfo.PageIndex && this.HeaderInfo.Count > 0 && CurRow >= this.HeaderInfo.Count)
//            this.HeaderInfo.PageIndex = CurPage;
//    Without it the header is drawn only on pages reached by running out of room.
//
// 2. A table never leaves only its header rows at the foot of its first page; Word
//    (2010 and later, the condition ONLYOFFICE's own rule for this uses) carries the
//    table over whole, and a heading kept with it goes too.
//    a. Forced keep: `if (0 === nKeepNextForceRow)` becomes
//           if (0 === nKeepNextForceRow || (nKeepNextForceRow <= this.HeaderInfo.Count && Word14+))
//    b. A first data row that cannot split and does not fit: before
//       private_RecalculatePageBottomBorders(CurPage),
//           if (bNextPage && 0 === CurPage && this.HeaderInfo.Count > 0 && Word14+
//               && this.RowsInfo[0].FirstPage && 0 <= this.Pages[0].LastRow && this.Pages[0].LastRow < this.HeaderInfo.Count
//               && docPrev && isNotFirstOnPage)
//           { this.Pages[0].LastRow = -1; this.Pages[0].LastRowSplit = false; this.HeaderInfo.PageIndex = -1; }
//       (ONLYOFFICE checks the header-only case before it moves a row that cannot
//       split, so the row it moves leaves the header behind.)
//
// Minified names: gu = HeaderInfo (Ck PageIndex, Rq Count), Ta = Pages (Lx LastRow,
// j0b LastRowSplit), Yd = RowsInfo (JO FirstPage), a = CurPage, Na = CurRow,
// g = FirstRow, Ga = nKeepNextForceRow, I = bNextPage, y = the logic document,
// AscCommon.QAa = document_compatibility_mode_Word14, ht = the element before,
// cI(...) = is first element on the page.
import { readFileSync, writeFileSync } from "node:fs";

const WORD14 = "(y&&y.XQ?y.XQ():AscCommon.h5c)>=AscCommon.QAa";
const FIRST_ON_PAGE = "cI(this.Ba,this.aQ(a),this.Pe())";

export const PATCHES: { name: string; find: string; replace: string }[] = [
  {
    name: "header repeats over rows a keep carried",
    find: "if(Ga>g&&Na===Ga){da=I=!0;n=Na-1;this.Ta[a].Lx=Na-1;this.Ta[a].j0b=!1;break}",
    replace: "if(Ga>g&&Na===Ga){da=I=!0;n=Na-1;this.Ta[a].Lx=Na-1;this.Ta[a].j0b=!1;-1===this.gu.Ck&&0<this.gu.Rq&&Na>=this.gu.Rq&&(this.gu.Ck=a);break}",
  },
  {
    name: "a keep never leaves the header alone",
    find: `if(0===Ga)if(0!==a||!this.ht||this.Ba&&${FIRST_ON_PAGE})Ga=-1;else return this.xKa(),this.i8a=!1,Sx`,
    replace: `if(0===Ga||Ga<=this.gu.Rq&&${WORD14})if(0!==a||!this.ht||this.Ba&&${FIRST_ON_PAGE})Ga=-1;else return this.xKa(),this.i8a=!1,Sx`,
  },
  {
    name: "a row that cannot split never leaves the header alone",
    find: "var Kb=this.Ta[a],Tb=Kb.Lx;Kb.B6=0;Kb.uKc=[];",
    replace: `I&&0===a&&0<this.gu.Rq&&${WORD14}&&this.Yd[0]&&this.Yd[0].JO&&0<=this.Ta[0].Lx&&this.Ta[0].Lx<this.gu.Rq&&this.ht&&this.Ba&&!${FIRST_ON_PAGE}&&(this.Ta[0].Lx=-1,this.Ta[0].j0b=!1,this.gu.Ck=-1);var Kb=this.Ta[a],Tb=Kb.Lx;Kb.B6=0;Kb.uKc=[];`,
  },
];

export function patchSdk(source: string): string {
  let out = source;
  for (const p of PATCHES) {
    const count = out.split(p.find).length - 1;
    if (count !== 1) throw new Error(`sdkjs patch "${p.name}": expected 1 match, found ${count}`);
    out = out.replace(p.find, () => p.replace);
  }
  return out;
}

if (import.meta.main) {
  const file = process.argv[2];
  if (!file) throw new Error("usage: bun sdkjs-patch.ts <sdk-all.js>");
  writeFileSync(file, patchSdk(readFileSync(file, "utf8")));
  console.log(`patched ${file}: ${PATCHES.map((p) => p.name).join("; ")}`);
}
