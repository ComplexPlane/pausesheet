import SS = GoogleAppsScript.Spreadsheet;

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

    range.setNumberFormat('@'); // '@' means plain text for some reason...
    // For more info: https://www.blackcj.com/blog/2015/05/18/cell-number-formatting-with-google-apps-script/

    setBgColor(range);
    setFontStuff(range);
    // setBorders(range); // TODO optimize
    doMerges(range);
};

function setFontStuff(range: SS.Range) {
    const fontSizes = genRangeArray(range);
    const fontColors = genRangeArray(range);
    const fontFamilies = genRangeArray(range);
    const fontStyles = genRangeArray(range);
    const fontWeights = genRangeArray(range);

    for (let row = 0; row < range.getNumRows(); row++) {
        for (let col = 0; col < range.getNumColumns(); col++) {
            fontColors[row][col] = 'black';
            fontStyles[row][col] = 'normal'; // TODO adjustment description should be italic

            if (col == 0) {
                // Stage name
                fontSizes[row][col] = 10;
                fontFamilies[row][col] = 'Arial';
                fontWeights[row][col] = 'bold';

            } else {
                // Frames
                if (col % 2 == 1) {
                    // Input column
                    fontSizes[row][col] = 15;
                    fontFamilies[row][col] = 'Calibri';
                    fontWeights[row][col] = 'bold';
                } else {
                    // Frame column
                    fontSizes[row][col] = 11;
                    fontFamilies[row][col] = 'arial';
                    fontWeights[row][col] = 'normal';
                }
            }
        }
    }

    range.setFontSizes(fontSizes);
    range.setFontColors(fontColors);
    range.setFontFamilies(fontFamilies);
    range.setFontStyles(fontStyles);
    range.setFontWeights(fontWeights);
}

function setBorders(range: SS.Range) {
    // Clear borders initially
    range.setBorder(false, false, false, false, false, false);

    const values = range.getValues();
    const nonEmptyCellLocs = [];
    for (let row = 1; row <= range.getNumRows(); row++) {
        for (let col = 1; col <= range.getNumColumns(); col++) {
            if (values[row - 1][col - 1] !== '') {
                const absRow = row + range.getRow() - 1;
                const absCol = col + range.getColumn() - 1;
                nonEmptyCellLocs.push(`R${absRow}C${absCol}`);
            }
        }
    }

    const rangeList = range.getSheet().getRangeList(nonEmptyCellLocs); // This is slow...
    rangeList.setBorder(true, true, true, true, true, true);
}

function setBgColor(range: SS.Range) {
    const backgrounds = genRangeArray(range);

    // Sample bg color
    const bgColor = range.getCell(1, 2).getBackground();
    for (let row = 0; row < range.getNumRows(); row++) {
        for (let col = 0; col < range.getNumColumns(); col++) {
            const cell = range.getCell(row + 1, col + 1);
            if (col == 0) {
                // Name of stage, preserve bg color
                backgrounds[row][col] = cell.getBackground();
            } else if (cell.isBlank()) {
                backgrounds[row][col] = '';
            } else {
                backgrounds[row][col] = bgColor;
            }
        }
    }

    range.setBackgrounds(backgrounds);
}

function getSubRange(range: SS.Range, row, col, numRows, numCols): SS.Range {
    const absRow = range.getRow() + row - 1;
    const absCol = range.getColumn() + col - 1;
    return range.getSheet().getRange(absRow, absCol, numRows, numCols);
}

function doMerges(range: SS.Range) {
    for (let row = 1; row < range.getNumRows(); row += 2) {
        for (let col = 1; col <= range.getNumColumns(); col++) {
            const below = getSubRange(range, row, col, 2, 1);
            below.merge();
        }
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

// // Returns the column of the next data, or null if the rest is blanks
// function findNextDataInRow(values: any[][], row: number, col: number): number | null {
//     for (let newCol = col + 1; newCol - 1 < values[row].length; newCol++) {
//         if (values[row - 1][newCol - 1] != '') {
//             return newCol;
//         }
//     }
//     return null;
// }

// // Returns the row of the next data, or null if the rest is blanks
// function findNextDataInCol(values: any[][], row: number, col: number): number | null {
//     for (let newRow = row + 1; newRow - 1 < values.length; newRow++) {
//         if (values[newRow - 1][col - 1] != '') {
//             return newRow;
//         }
//     }
//     return null;
// }