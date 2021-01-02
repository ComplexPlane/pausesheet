// TODO
// - Fix merge-over-merge bug
// - Italicize adjustment
// - Normalize frame numbers / ranges (~ symbol, trailing 0)

import SS = GoogleAppsScript.Spreadsheet;

const frameRangeRe = /^((\d?\d)\.\d\d)( ?[～\-\~] ?((\d?\d\.)?\d\d))?$/g;
// Group 1 is start frame, group 4 is end frame (if it exists)

/* Test cases:
55.21
50.83
54.30～26
54.76-75
54.76~75
54.30 ～ 26
54.76 - 75
54.76 ~ 75
55.00-54.98
*/

interface Context {
    range: SS.Range;
    values: any[][];
    bgColors: any[][];
    fontSizes: any[][];
    fontFamilies: any[][];
    fontStyles: any[][];
    fontWeights: any[][];
    horizontalAlignments: any[][];
    verticalAlignments: any[][];
    mergedCells: boolean[][];
}

function onOpen() {
    const ui = SpreadsheetApp.getUi();
    // Or DocumentApp or FormApp.
    ui.createMenu('Scripts')
        .addItem('Format stage table', 'formatStageTable')
        .addToUi();
}

// From: https://stackoverflow.com/questions/15673038/how-do-i-copy-a-row-with-both-values-and-formulas-to-an-array
// needed to preserve hyperlinks
function getValuesAndFormulas(range: SS.Range): any[][] {
    const formulas = range.getFormulas();
    const values = range.getValues();
    const merge = new Array(formulas.length);
    for (let i in formulas) {
        merge[i] = new Array(formulas[i].length);
        for (let j in formulas[i]) {
            merge[i][j] = formulas[i][j] !== '' ? formulas[i][j] : values[i][j];
        }
    }

    return merge;
}

function genRangeArray<T>(range: SS.Range, defaultValue: T): T[][] {
    const arr: any[][] = [];
    for (let row = 1; row <= range.getNumRows(); row++) {
        const rowArr: T[] = [];
        for (let col = 1; col <= range.getNumColumns(); col++) {
            rowArr.push(defaultValue);
        }
        arr.push(rowArr);
    }
    return arr;
}

function findMergedCells(range: SS.Range): boolean[][] {
    const mergedRanges = range.getMergedRanges();
    const bools = genRangeArray(range, false);

    for (let mergedRange of mergedRanges) {
        for (let absRow = mergedRange.getRow(); absRow <= mergedRange.getLastRow(); absRow++) {
            for (let absCol = mergedRange.getColumn(); absCol <= mergedRange.getLastColumn(); absCol++) {
                const row = absRow - range.getRow() + 1;
                const col = absCol - range.getColumn() + 1;
                bools[row - 1][col - 1] = true;
            }
        }
    }

    return bools;
}

function formatStageTable() {
    const sheet = SpreadsheetApp.getActive();
    const range = sheet.getActiveSheet().getActiveRange();
    const ctx: Context = {
        range: range,
        values: getValuesAndFormulas(range),
        bgColors: range.getBackgrounds(),
        fontSizes: range.getFontSizes(),
        fontFamilies: range.getFontFamilies(),
        fontStyles: range.getFontStyles(),
        fontWeights: range.getFontWeights(),
        horizontalAlignments: range.getHorizontalAlignments(),
        verticalAlignments: range.getVerticalAlignments(),
        mergedCells: findMergedCells(range),
    };
    // After this point, zero additional reads from the sheet are allowed

    doMerges(ctx);
    setBgColor(ctx);
    setFontStuff(ctx);
    setBorders(ctx);
    rewriteArrows(ctx);
    // normalizeFrameRanges(ctx); // TODO fix (only works like 2/3 of the time??)

    range.setNumberFormat('@'); // '@' means plain text for some reason...
    range.setValues(ctx.values);
    // For more info: https://www.blackcj.com/blog/2015/05/18/cell-number-formatting-with-google-apps-script/
};

function setFontStuff(ctx: Context) {
    for (let row = 0; row < ctx.range.getNumRows(); row++) {
        for (let col = 0; col < ctx.range.getNumColumns(); col++) {
            ctx.fontStyles[row][col] = 'normal'; // TODO adjustment description should be italic
            ctx.horizontalAlignments[row][col] = 'center';
            ctx.verticalAlignments[row][col] = 'middle';

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
    ctx.range.setFontFamilies(ctx.fontFamilies);
    ctx.range.setFontStyles(ctx.fontStyles);
    ctx.range.setFontWeights(ctx.fontWeights);
    ctx.range.setHorizontalAlignments(ctx.horizontalAlignments);
    ctx.range.setVerticalAlignments(ctx.verticalAlignments);
}

function setBorders(ctx: Context) {
    // Clear borders initially
    ctx.range.setBorder(false, false, false, false, false, false);

    for (let row = 1; row <= ctx.range.getNumRows(); row++) {
        for (let col = 1; col <= ctx.range.getNumColumns(); col++) {
            if (ctx.values[row - 1][col - 1] !== '') {
                ctx.range.getCell(row, col).setBorder(true, true, true, true, true, true);
            }
        }
    }
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

function emptiesRightOfCell(ctx: Context, row: number, col: number): boolean {
    for (let newCol = col + 1; newCol <= ctx.range.getNumColumns(); newCol++) {
        if (ctx.values[row - 1][newCol - 1] !== '') return false;
    }
    return true;
}

function expandDown(ctx: Context, row: number, col: number): SS.Range {
    let newRow = row + 1;
    for (; newRow <= ctx.range.getNumRows(); newRow++) {
        // Stop when...
        if (ctx.values[newRow - 1][col - 1] !== '' // Cell is empty
            || ctx.mergedCells[newRow - 1][col - 1] // Cell is already part of a merge
            || emptiesRightOfCell(ctx, newRow, col)) break; // There are no non-empty cells to the right (end of tree)
    }

    return getSubRange(ctx.range, row, col, newRow - row, 1);
}

function doMerges(ctx: Context) {
    for (let col = 1; col < ctx.range.getNumColumns(); col++) {
        let row = 1;
        while (row <= ctx.range.getNumRows()) {
            const mergeRegion = expandDown(ctx, row, col);
            mergeRegion.merge();
            row += mergeRegion.getNumRows();
        }
    }
}

function rewriteArrows(ctx: Context) {
    for (let row = 1; row <= ctx.range.getNumRows(); row++) {
        for (let col = 2; col <= ctx.range.getNumColumns(); col += 2) {
            const text = String(ctx.values[row - 1][col - 1]).toLowerCase();
            switch (text) {
                case 'upleft':
                case 'leftup':
                case 'ul': {
                    ctx.values[row - 1][col - 1] = '↖';
                    break;
                };

                case 'upright':
                case 'rightup':
                case 'ur': {
                    ctx.values[row - 1][col - 1] = '↗';
                    break;
                }

                case 'downleft':
                case 'leftdown':
                case 'dl': {
                    ctx.values[row - 1][col - 1] = '↙';
                    break;
                }

                case 'downright':
                case 'rightdown':
                case 'dr': {
                    ctx.values[row - 1][col - 1] = '↘';
                    break;
                }

                case 'left':
                case 'l': {
                    ctx.values[row - 1][col - 1] = '←';
                    break;
                }

                case 'right':
                case 'r': {
                    ctx.values[row - 1][col - 1] = '→';
                    break;
                }

                case 'up':
                case 'u': {
                    ctx.values[row - 1][col - 1] = '↑';
                    break;
                }

                case 'down':
                case 'd': {
                    ctx.values[row - 1][col - 1] = '↓';
                    break;
                }

                case 'neutral':
                case 'n': {
                    ctx.values[row - 1][col - 1] = 'N';
                    break;
                }
            }
        }
    }
}

function normalizeFrameRanges(ctx: Context) {
    for (let row = 1; row <= ctx.range.getNumRows(); row++) {
        for (let col = 3; col <= ctx.range.getNumColumns(); col += 2) {
            const frameText = String(ctx.values[row - 1][col - 1]);
            const match = frameRangeRe.exec(frameText);
            if (match === null) continue;

            Logger.log(match);

            let normalized = '';
            if (match[3] !== undefined) {
                // Frame range
                // normalized = `${match[1]}~${match[4]}`;
                normalized = 'range';
            } else {
                // Single frame
                // normalized = match[1];
                normalized = 'single';
            }

            ctx.values[row - 1][col - 1] = normalized;
        }
    }
}