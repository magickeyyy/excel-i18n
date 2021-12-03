const fs = require("fs");
const path = require("path");
const { pickSheet } = require("../index");
const { expect } = require("chai");

describe("excel-i18n", function () {
    it("构建产物", function () {
        try {
            fs.rmdirSync(path.resolve(process.cwd(), "test/locale"), {
                recursive: true,
            });
        } catch (e) {}
        pickSheet({
            inputPath: "test/i18n.xlsx",
            outputDir: "test/locale",
            extension: ".json",
            sheetName: "宝箱",
            keyX: "A",
            keyY1: 5,
            keyY2: 14,
            titleX: ["D", "E"],
            titleY: 3,
            handleTitle: (_, content) => {
                if (content?.split("/")[1]) return content?.split("/")[1];
                console.log(content, _, "标题不合规范");
                process.exit();
            },
            handleContent(_, content) {
                if (
                    _ === "D10" &&
                    (!content ||
                        !content.includes("{nickname}") ||
                        !content.includes("{cash}"))
                ) {
                    console.log(content, "占位符错误");
                    process.exit();
                }
                return content;
            },
        });
        const locales = fs.readdirSync("./test/locale", {
            withFileTypes: true,
        });
        const names = locales.map((v) => v.name).join();
        expect(names, "文件名是否正确").to.equal("en-US.json,zh-CN.json");
        expect(locales.length, "文件数量是否正确").to.equal(2);
        const keys = [
            "TREASURE_CHEST",
            "CHEST_OPEN",
            "SPECIAL_REWARD1",
            "SPECIAL_REWARD",
            "SPECIAL_CHEST",
            "KGHH",
            "REWARD_YOUR",
            "REWARD_SHARE_DESC_1",
            "ONECE",
            "RECORD_HELP",
        ];
        const contents = [
            fs.readFileSync("./test/locale/zh-CN.json", {
                encoding: "utf-8",
            }),
            fs.readFileSync("./test/locale/en-US.json", {
                encoding: "utf-8",
            }),
        ].map((v) => Object.keys(JSON.parse(v || {})));
        const keysFine = contents
            .map((v) => v.join())
            .every((v) => v === keys.join());
        const contentsFine = contents
            .map((v) => Object.values(v).every((v) => v))
            .every((v) => v);
        expect(keysFine, "多语言KEY是否完整正确").to.equal(true);
        expect(contentsFine, "多语言每项是否有类容").to.equal(true);
    });
});
