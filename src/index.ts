import "colors";
import path from "path";
import fs from "fs";
import { readFile } from "xlsx";
import type { CellObject } from "xlsx";

const pwd = process.cwd();
const Letters = String.fromCharCode(
    ...Array.from({ length: 26 }).map((_, i) => 65 + i)
).split("");
function getCellName(count: number) {
    if (count <= 26) {
        return Letters[count - 1];
    }
    let arr: number[] = [];
    while (count > 26) {
        let data = Math.floor(count / 26);
        arr.unshift(count - data * 26);
        count = data;
        if (count < 26) {
            arr.unshift(count - 1);
        }
    }
    return arr.map((v) => Letters[v]).join("");
}
/*
目前发现wps开过的文档会增加默认一个隐藏页：WpsReserved_CellImgList,所以取sheet还是以sheet名为准
取值第一行是表头信息（region[0]），第一列是key信息
坐标位置都是从1开始
*/
function fn({
    filePath,
    sheetName,
    region,
    colmuns,
    sheetIndex = 1,
    keyIndex = 1,
    titleIndex = 1,
    formateLang,
}: {
    filePath: string;
    sheetIndex?: number;
    sheetName?: string;
    region: [number, number, number, number]; // 左上角行数，左上角列数，右上角列数，右下角行数，从1开始
    colmuns: [number, number]; // 取多语言的开始列数和结束列数,从1开始，连续
    keyIndex?: number; // 从region[1]开始，第n列是key列，从1开始
    titleIndex?: number; // 从region[0]开始第n行是标题栏，从1开始
    formateLang: (title: string) => string;
}) {
    const workbook = readFile(filePath);
    const targetSheetName = sheetName || workbook.SheetNames[sheetIndex];
    const [lx, ly, rx, ry] = region;
    const curSheet = workbook.Sheets[targetSheetName];
    let cy = ly + 1;
    const keys: string[] = [];
    while (cy <= ry) {
        const keyx = cy + keyIndex - 1;
        const curKey = curSheet[Letters[lx - 1] + keyx]?.v?.trim?.();
        if (curKey && !keys.includes(curKey)) {
            keys.push(curKey);
        } else {
            console.log(curSheet[Letters[lx - 1] + keyx], "不存在".red);
            process.exit();
        }
        ++cy;
    }
    console.log("keys", keys);
    const [ls, le] = colmuns;
    const langs: string[] = [];
    let clx = region[1] + titleIndex - 1;
    let langStart = ls;
    while (langStart <= le) {
        const title = curSheet[Letters[langStart - 1] + clx]?.v?.trim();
        const lang = title ? formateLang(title) : "";
        if (title && lang && !langs.includes(lang)) {
            langs.push(lang);
        } else {
            console.log("语言项提取结果错误".red, lang);
        }
        ++langStart;
    }
    const data: Record<string, Record<string, string>> = {};
    for (let i = ls; i <= le; i++) {
        const lang = langs[i - ls];
        if (!(lang in data)) {
            data[lang] = {};
            for (let j = lx + 1; j <= ry; j++) {
                data[lang][keys[j - lx - 1]] =
                    curSheet[getCellName(i) + j]?.v?.trim() || "取错了";
            }
        }
    }
    fs.writeFileSync(path.resolve(pwd, "tt.json"), JSON.stringify(data));
    process.exit();
}
fn({
    filePath: path.resolve(pwd, "excel", "Clap house 文本表.xlsx"),
    sheetIndex: 7,
    sheetName: "PDD活动宝箱文本",
    region: [1, 4, 8, 14],
    colmuns: [5, 8],
    formateLang: (title) => title.split("/")[1],
});
