import "colors";
import path from "path";
import fs from "fs";
import XLSX from "xlsx/";

const pwd = process.cwd();
/*
可能会有隐藏页,所以取sheet还是以sheetName名为准
坐标位置都是从1开始，[列数,行数]
多语言存放区域（SCOPE），起点列数应该和标题栏起点列数相等，终点列数应该和标题栏终点列数相等；起点行数应该和KEY起点行数相等，终点行数应该和KEY列终点行数相等
*/
export function pickSheet({
    inputPath,
    outputDir,
    extension,
    sheetName,
    sheetIndex = 0,
    keyX,
    keyY1,
    keyY2,
    titleX,
    titleY,
    handleTitle,
    handleContent,
    handleKey,
    callback,
}: {
    inputPath: string; // 文件在process.pwd()下的路径
    outputDir: string; // 输出文件夹
    extension: ".json" | ".ts" | ".js"; // 输出文件后缀名
    sheetIndex?: number;
    sheetName?: string;
    keyX: string; // KEY列序号
    keyY1: number; // KEY起点行号
    keyY2: number; // KEY终点行号
    titleX: string[]; // title每列序号
    titleY: number; // title在哪一行
    handleKey?: (name: string, content?: string) => string | void;
    handleTitle: (name: string, content?: string) => string | void;
    handleContent?: (name: string, content?: string) => string | void;
    callback?: () => void;
}) {
    const workbook = XLSX.readFile(path.resolve(pwd, inputPath));
    const targetSheetName = sheetName || workbook.SheetNames[sheetIndex];
    const curSheet = workbook.Sheets[targetSheetName];
    if (
        !titleX.length ||
        titleX.concat(keyX).find((v) => !/^[A-Z]+$/g.test(v))
    ) {
        console.log("keyX和titleX每项都应该是大写字母".red);
        process.exit();
    }
    if (keyY1 > keyY2) {
        console.log("keyY2应该大于keyY1".red);
        process.exit();
    }
    // 处理KEY列
    const keys: string[] = [];
    Array.from({ length: keyY2 - keyY1 + 1 }).map((_, i) => {
        const v = curSheet[keyX + (keyY1 + i)]?.v;
        const key = handleKey?.(keyX + (keyY1 + i), v) || v?.trim();
        if (key && !keys.includes(key)) {
            keys.push(key);
        } else {
            console.log(curSheet[keyX + (keyY1 + i)]?.v, "KEY错误或者重复".red);
            process.exit();
        }
    });
    const langs: string[] = [];
    titleX.map((x) => {
        const v = curSheet[x + titleY]?.v;
        const lang = handleTitle(x + titleY, v);
        if (lang && !langs.includes(lang)) {
            langs.push(lang);
        } else {
            console.log(v, lang, "语言错误或重复".red);
            process.exit();
        }
    });
    const locale: Record<string, Record<string, string>> = {};
    for (let i = 0; i < titleX.length; i++) {
        const lang = langs[i];
        if (!locale.hasOwnProperty(lang)) {
            locale[lang] = {};
            for (let j = keyY1; j <= keyY2; j++) {
                const v = curSheet[titleX[i] + j]?.v;
                const hv =
                    handleContent?.(
                        titleX[i] + j,
                        curSheet[titleX[i] + j]?.v
                    ) || curSheet[titleX[i] + j]?.v?.trim();
                if (!hv) {
                    console.log(`${titleX[i] + j}没内容`.red);
                }
                locale[lang][keys[j - keyY1]] = hv;
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
                ? JSON.stringify(data, null, "    ")
                : "export default " + JSON.stringify(data, null, "    ")
        )
    );
    console.log("输出结果：".green, path.resolve(pwd, outputDir).bgBlue);
    callback?.();
}
