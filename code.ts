import SS = GoogleAppsScript.Spreadsheet;

interface Context {
    range: SS.Range;
    values: any[][];
    bgColors: any[][];
    fontSizes: any[][];
    fontColors: any[][];
    fontFamilies: any[][];
    fontStyles: any[][];
    fontWeights: any[][];
}

function onOpen() {
    const ui = SpreadsheetApp.getUi();
    // Or DocumentApp or FormApp.
    ui.createMenu('Scripts')
        .addItem('Format Stage Table', 'formatStageTable')
        .addToUi();
}

function genRangeArray(range: SS.Range): any[][] {
    const arr: any[][] = [];
    for (let row = 1; row <= range.getNumRows(); row++) {
        arr.push(new Array(range.getNumColumns()));
    }
    return arr;
}

function formatStageTable() {
    const sheet = SpreadsheetApp.getActive();
    const range = sheet.getActiveSheet().getActiveRange();
    const ctx: Context = { 
        range: range,
        values: range.getValues(),
        bgColors: range.getBackgrounds(),
        fontSizes: range.getFontSizes(),
        fontColors: range.getFontColors(),
        fontFamilies: range.getFontFamilies(),
        fontStyles: range.getFontStyles(),
        fontWeights: range.getFontWeights(),
    };
    // After this point, zero additional reads from the sheet are allowed

    doMerges(ctx);
    setBgColor(ctx);
    setFontStuff(ctx);
    // setBorders(range); // TODO optimize

    range.setNumberFormat('@'); // '@' means plain text for some reason...
    // For more info: https://www.blackcj.com/blog/2015/05/18/cell-number-formatting-with-google-apps-script/
};

function setFontStuff(ctx: Context) {
    for (let row = 0; row < ctx.range.getNumRows(); row++) {
        for (let col = 0; col < ctx.range.getNumColumns(); col++) {
            ctx.fontColors[row][col] = 'black';
            ctx.fontStyles[row][col] = 'normal'; // TODO adjustment description should be italic

            if (col == 0) {
                // Stage name
                ctx.fontSizes[row][col] = 10;
                ctx.fontFamilies[row][col] = 'Arial';
                ctx.fontWeights[row][col] = 'bold';

            } else {
                // Frames
                if (col % 2 == 1) {
                    // Input column
                    ctx.fontSizes[row][col] = 15;
                    ctx.fontFamilies[row][col] = 'Calibri';
                    ctx.fontWeights[row][col] = 'bold';
                } else {
                    // Frame column
                    ctx.fontSizes[row][col] = 11;
                    ctx.fontFamilies[row][col] = 'arial';
                    ctx.fontWeights[row][col] = 'normal';
                }
            }
        }
    }

    ctx.range.setFontSizes(ctx.fontSizes);
    ctx.range.setFontColors(ctx.fontColors);
    ctx.range.setFontFamilies(ctx.fontFamilies);
    ctx.range.setFontStyles(ctx.fontStyles);
    ctx.range.setFontWeights(ctx.fontWeights);
}

function setBorders(ctx: Context) {
    // Clear borders initially
    ctx.range.setBorder(false, false, false, false, false, false);

    const nonEmptyCellLocs = [];
    for (let row = 1; row <= ctx.range.getNumRows(); row++) {
        for (let col = 1; col <= ctx.range.getNumColumns(); col++) {
            if (ctx.values[row - 1][col - 1] !== '') {
                const absRow = row + ctx.range.getRow() - 1;
                const absCol = col + ctx.range.getColumn() - 1;
                nonEmptyCellLocs.push(`R${absRow}C${absCol}`);
            }
        }
    }

    const rangeList = ctx.range.getSheet().getRangeList(nonEmptyCellLocs); // This is slow... should this be done during "read" phase?
    rangeList.setBorder(true, true, true, true, true, true);
}

function setBgColor(ctx: Context) {
    // Sample bg color
    const bgColor = ctx.bgColors[0][1];
    for (let row = 0; row < ctx.range.getNumRows(); row++) {
        for (let col = 0; col < ctx.range.getNumColumns(); col++) {
            // Preserve color of first column (name of stage)
            if (col != 0) {
                if (ctx.values[row][col] === '') {
                    ctx.bgColors[row][col] = '';
                } else {
                    ctx.bgColors[row][col] = bgColor;
                }
            }
        }
    }

    ctx.range.setBackgrounds(ctx.bgColors);
}

function getSubRange(range: SS.Range, row, col, numRows, numCols): SS.Range {
    const absRow = range.getRow() + row - 1;
    const absCol = range.getColumn() + col - 1;
    return range.getSheet().getRange(absRow, absCol, numRows, numCols);
}

function expandRight(ctx: Context, row: number, col: number): SS.Range {
    let newCol = col + 1;
    for (; newCol <= ctx.range.getNumColumns(); newCol++) {
        if (ctx.values[row - 1][newCol - 1] !== '') break;
    }

    return getSubRange(ctx.range, row, col, 1, newCol - col);
}

function expandDown(ctx: Context, row: number, col: number): SS.Range {
    let newRow = row + 1;
    for (; newRow <= ctx.range.getNumRows(); newRow++) {
        if (ctx.values[newRow - 1][col - 1] !== '') break;
    }

    return getSubRange(ctx.range, row, col, newRow - row, 1);
}

function doMerges(ctx: Context) {
    for (let col = 1; col < ctx.range.getNumColumns(); col++) {
        const mergeRegion = expandDown(ctx, 1, col);
        mergeRegion.merge();
    }

    // const values = range.getValues();

    // for (let col = 1; col <= range.getNumColumns(); col++) {
    //     let row = 1;
    //     while (row <= range.getNumRows()) {
    //         let newRow = findNextDataInCol(values, row, col);
    //         if (newRow === null) newRow = range.getNumColumns();

    //         const mergeRange = range.getSheet().getRange(range.getRow() + row - 1, col, row, 1)
    //     }
    // }
}