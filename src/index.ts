import "colors";
import path from "path";
import fs from "fs";
import { readFile } from "xlsx";

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
目前发现wps开过的文档会增加默认一个隐藏页：WpsReserved_CellImgList,所以取sheet还是以sheetName名为准
坐标位置都是从1开始，[列数,行数]
多语言存放区域（SCOPE），起点列数应该和标题栏起点列数相等，终点列数应该和标题栏终点列数相等；起点行数应该和KEY起点行数相等，终点行数应该和KEY列终点行数相等
*/
function fn({
    inputPath,
    outputDir,
    extension,
    sheetName,
    sheetIndex = 0,
    TITLE,
    KEY,
    SCOPE,
}: {
    inputPath: string; // 文件在process.pwd()下的路径
    outputDir: string; // 输出文件夹
    extension: ".json" | ".ts" | ".js"; // 输出文件后缀名
    sheetIndex?: number;
    sheetName?: string;
    // 标题栏，一般包含多语言信息，最终会以语言名做文件名
    TITLE: {
        start: [number, number];
        end: [number, number];
        handle: (name: string, content?: string) => string | void;
    };
    // KEY列，不包含KEY表头
    KEY: {
        start: [number, number];
        end: [number, number];
        handle?: (name: string, content?: string) => string | void;
    };
    // 多语言在表格中应该是一个矩形，不能隔断
    SCOPE: {
        leftTop: [number, number];
        rightBottom: [number, number];
        handle?: (name: string, content?: string) => string | void;
    };
}) {
    const workbook = readFile(path.resolve(pwd, inputPath));
    const targetSheetName = sheetName || workbook.SheetNames[sheetIndex];
    const curSheet = workbook.Sheets[targetSheetName];
    const { start: keyStart, end: keyEnd, handle: keyHandle } = KEY;
    const { start: titleStart, end: titleEnd, handle: titleHandle } = TITLE;
    const { leftTop, rightBottom, handle: contentHandle } = SCOPE;
    if (
        keyStart[0] < 1 ||
        keyEnd[0] < 1 ||
        keyStart[1] < 2 ||
        keyEnd[1] < 2 ||
        keyStart[0] !== keyEnd[0] ||
        keyStart[1] >= keyEnd[1]
    ) {
        console.log("KEY参数错误".red);
        process.exit();
    }
    if (
        titleStart[0] < 1 ||
        titleEnd[0] < 1 ||
        titleStart[1] < 1 ||
        titleEnd[1] < 1 ||
        titleStart[1] !== titleEnd[1] ||
        titleStart[0] >= titleEnd[0]
    ) {
        console.log("TITLE参数错误".red);
        process.exit();
    }
    if (
        leftTop[0] !== titleStart[0] ||
        leftTop[1] !== keyStart[1] ||
        rightBottom[0] !== titleEnd[0] ||
        rightBottom[1] !== keyEnd[1]
    ) {
        console.log("SCOPE参数错误".red);
        process.exit();
    }
    // 处理KEY列
    const keys: string[] = [];
    const keyLetter = Letters[keyStart[0] - 1];
    let keyY = keyStart[1];
    while (keyY <= keyEnd[1]) {
        const curKey =
            keyHandle?.(keyLetter + keyY, curSheet[keyLetter + keyY]?.v) ||
            curSheet[keyLetter + keyY]?.v?.trim?.();
        if (curKey && !keys.includes(curKey)) {
            keys.push(curKey);
        } else {
            console.log(curKey, "KEY错误或者重复".red);
            process.exit();
        }
        ++keyY;
    }
    const langs: string[] = [];
    let langX = titleStart[0];
    while (langX <= titleEnd[0]) {
        const lang = titleHandle(
            getCellName(langX) + titleStart[1],
            curSheet[getCellName(langX) + titleStart[1]]?.v
        );
        if (lang && !langs.includes(lang)) {
            langs.push(lang);
            ++langX;
        } else {
            console.log("从标题栏提取语言错误或者语言重复".red, lang);
            process.exit();
        }
    }
    const locale: Record<string, Record<string, string>> = {};
    for (let i = leftTop[0]; i <= rightBottom[0]; i++) {
        const lang = langs[i - leftTop[0]];
        if (!locale.hasOwnProperty(lang)) {
            locale[lang] = {};
            for (let j = leftTop[1]; j <= rightBottom[1]; j++) {
                const content = curSheet[getCellName(i) + j]?.v;
                if (content == undefined) {
                    console.log(`${getCellName(i) + j}没内容`.red);
                }
                locale[lang][keys[j - leftTop[1]]] =
                    contentHandle?.(getCellName(i) + j, content) ||
                    content?.trim();
            }
        }
    }
    try {
        fs.readdirSync(path.resolve(pwd, outputDir));
    } catch (e) {
        fs.mkdirSync(path.resolve(pwd, outputDir));
    }
    Object.entries(locale).map(([filename, data]) =>
        fs.writeFileSync(
            path.resolve(pwd, outputDir, filename + extension),
            extension === ".json"
                ? JSON.stringify(data, null, "\t")
                : "export default " + JSON.stringify(data, null, "\t")
        )
    );
    console.log("输出结果：".green, path.resolve(pwd, outputDir).bgBlue);
    process.exit();
}
fn({
    inputPath: "excel/Clap house 文本表.xlsx",
    outputDir: "locale",
    extension: ".ts",
    sheetIndex: 7,
    sheetName: "PDD活动宝箱文本",
    TITLE: {
        start: [5, 4],
        end: [8, 4],
        handle: (_, content) => {
            if (content?.split("/")[1]) return content?.split("/")[1];
            console.log(content, "标题不合规范".red);
            process.exit();
        },
    },
    KEY: { start: [1, 5], end: [1, 14] },
    SCOPE: {
        leftTop: [5, 5],
        rightBottom: [8, 14],
    },
});
