
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

    public logDuplicates(rangeToCheck: string) {

        Excel.run(async (context) => {
            // Get the values in the range
            const cellRange = this._activeSheet.getRange(rangeToCheck).load("values");;

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
                            // Removes any color if the cell was red
                            cellRange.getCell(i, z).format.fill.clear();
                        }
                    }
                }
                // Log the coordinates of the duplicates, coords are relative
                // to the range not the sheet
                console.log(duplicateCoords.toString());
            });

        });
    }

    public emptyPromise(): Promise<string> {
        return new Promise((resolve, reject) => {
            resolve('asdf');
        });
    }
}
