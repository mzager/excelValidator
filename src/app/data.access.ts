
export class DataAccess {

    private _activeSheet: Excel.Worksheet;

    public setActiveSheet(name: string): Promise<Excel.Worksheet> {
        return new Promise((resolve, reject) => {
            this.getSheet(name).then(sheet => {
                this._activeSheet = sheet;
                resolve(sheet);
            });
        });
    }

    public getSheet(name: string): Promise<Excel.Worksheet> {
        // Function Returns a promise of a value in the future
        return new Promise((resolve, reject) => {
            // Go fetch the worksheet... this won't happen immediately...
            Excel.run(async context => {

                // Create a que of commands to run when we call sync...
                const ws = context.workbook.worksheets.getItem(name);

                // Request the worksheet "then" wait for it...
                context.sync().then(() => {

                    // Cool got the worksheet set to the ws variable...
                    // Resolve your promise
                    resolve(ws);
                });
            });
        });
    }

    // The value returned could be a string, number, or boolean depending on the
    // cell type
    public getValue(row: number, col: number): Promise<any> {
        return new Promise((resolve, reject) => {
            Excel.run(async context => {
                // Create queue of commands to get the value

                const cellRange = this._activeSheet.getCell(row, col).load('values');

                // Request the value "then" wait for it...
                context.sync().then(() => {
                    const cellValue = cellRange.values[0][0];

                    // Got the value, now resolve
                    resolve(cellValue);
                });
            });
        });
    }

    // Sets the value of a specified cell
    public setValue(row: number, col: number, value: any) {
        Excel.run(async context => {

            const cellRange = this._activeSheet.getCell(row, col).load('values');

            context.sync().then(() => {
                cellRange.values[0][0] = value;
            });

        });
    }

    public getColumnValues(col: string): Promise<Array<any>> {
        return new Promise((resolve, reject) => {
            Excel.run(async context => {
                // Create queue of commands to get the value

                const cellRange = this._activeSheet.getRange(col + ":" + col).load('values');
                const conditionalFormat = cellRange.conditionalFormats
                    .add(Excel.ConditionalFormatType.containsText);
                conditionalFormat.textComparison
                // Request the value "then" wait for it...
                context.sync().then(() => {
                    const colValues = cellRange.values;

                    // Got the value, now resolve
                    resolve(colValues);
                });
            });
        });
    }

    public logDuplicates(rangeToCheck: string): Promise<void> {
        return new Promise((resolve, reject) => {
            Excel.run(async (context) => {
                // Get the values in the range
                const cellRange = this._activeSheet.getRange(rangeToCheck).load("values");

                await context.sync().then(() => {
                    // Create a set to store unique vals, array for coords, and get range values
                    const duplicateValues = new Set();
                    const duplicateCoords = new Array();
                    const vals = cellRange.values;
                    // Iterate and check for duplicates
                    for (var i = 0; i < vals.length; i++) {
                        for (var z = 0; z < vals[i].length; z++) {
                            if (duplicateValues.has(vals[i][z])) {
                                // If it's a duplicate, add the cell address with the form
                                // (i, z) relative to the range, not the sheet
                                duplicateCoords.push("(" + i + ", " + z + ")");
                                // Turns the cell red
                                cellRange.getCell(i, z).format.fill.color = "#ff0000";
                            } else {
                                duplicateValues.add(vals[i][z]);
                                // Removes any color if the cell was reds
                                cellRange.getCell(i, z).format.fill.clear();
                            }
                        }
                    }
                    // Log the coordinates of the duplicates, coords are relative
                    // to the range not the sheet
                    console.log(duplicateCoords.toString());
                    resolve();
                });

            });
        });
    }

    public logNotInBaseline(baselineRange: string, toCompareRange: string): Promise<void> {
        return new Promise((resolve, reject) => {
            Excel.run(async (context) => {
                // Get the given range from the first sheet (Baseline data)
                const sheet = context.workbook.worksheets.getItem("Sheet1");
                const range = sheet.getRange(baselineRange).load("values");

                // Get the range from the sheet to be compared against the baseline data
                const sheet2 = context.workbook.worksheets.getItem("Sheet2");
                const range2 = sheet2.getRange(toCompareRange).load("values");

                // Creates the two sets to compare against each other
                var baselineValues = new Set();
                var valsToCompare = new Set();

                await context.sync().then(() => {
                    // Collect the unique values in the given range of the baseline data
                    const vals = range.values;
                    for (var i = 0; i < vals.length; i++) {
                        if (!baselineValues.has(vals[i])) {
                            baselineValues.add(vals[i] + "");
                        }
                    }
                    baselineValues.delete("");
                }).then(() => {
                    // Collect the unique values in the given range of the event data
                    const vals = range2.values;
                    for (var i = 0; i < vals.length; i++) {
                        if (!valsToCompare.has(vals[i])) {
                            valsToCompare.add(vals[i] + "");
                        }
                    }
                    valsToCompare.delete("");
                }).then(() => {
                    // Compares the two sets and logs any entries that are present in 
                    // the event data but not in baseline
                    let b = Array.from(valsToCompare);
                    for (var i = 0; i < b.length; i++) {
                        if (!valsToCompare.has(b[i] + "")) {
                            console.log(b[i] + "");
                        }
                    }
                });


            });
        });
    }

    public emptyPromise(): Promise<string> {
        return new Promise((resolve, reject) => {
            resolve('asdf');
        });
    }
}
