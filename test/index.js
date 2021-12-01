const pickSheet = require("./excel-i18n").pickSheet;

pickSheet({
    inputPath: "test/i18n.xlsx",
    outputDir: "test/locale",
    extension: ".ts",
    sheetName: "宝箱",
    keyX: "A",
    keyY1: 5,
    keyY2: 14,
    titleX: ["D", "E"],
    titleY: 4,
    handleTitle: (_, content) => {
        if (content?.split("/")[1]) return content?.split("/")[1];
        console.log(content, "标题不合规范");
        process.exit();
    },
});
