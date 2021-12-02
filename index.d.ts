import "colors";
export declare function pickSheet({
    inputPath,
    outputDir,
    extension,
    sheetName,
    sheetIndex,
    keyX,
    keyY1,
    keyY2,
    titleX,
    titleY,
    handleTitle,
    handleContent,
    handleKey,
}: {
    inputPath: string;
    outputDir: string;
    extension: ".json" | ".ts" | ".js";
    sheetIndex?: number;
    sheetName?: string;
    keyX: string;
    keyY1: number;
    keyY2: number;
    titleX: string[];
    titleY: number;
    handleKey?: (name: string, content?: string) => string | void;
    handleTitle: (name: string, content?: string) => string | void;
    handleContent?: (name: string, content?: string) => string | void;
}): void;
