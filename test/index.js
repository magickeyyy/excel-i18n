const pickSheet = require("../dist/index").pickSheet;

pickSheet({
    inputPath: "test/excel/Clap house 文本表.xlsx",
    outputDir: "test/locale",
    extension: ".ts",
    sheetIndex: 7,
    sheetName: "PDD活动宝箱文本",
    keyX: "A",
    keyY1: 5,
    keyY2: 14,
    titleX: ["E", "F", "G", "H"],
    titleY: 4,
    handleTitle: (_, content) => {
        if (content?.split("/")[1]) return content?.split("/")[1];
        console.log(content, "标题不合规范");
        process.exit();
    },
});
