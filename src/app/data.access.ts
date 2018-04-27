export class DataAccess {

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
    public getValue(row: number, col: number): Promise<string | number | boolean> {
        return new Promise((resolve, reject) => {
            Excel.run(async context => {
                // Create queue of commands to get the value                

                // const cellRange = thisSheet.getRange(row, col).load("values");

                // Request the value "then" wait for it...
                context.sync().then(() => {
                    // const cellValue = cellRange.values[0][0];

                    // Got the value, now resolve
                    // resolve(cellValue);
                });
            });
        });
    }

    public getColumnValues(col: string): Promise<Array<string | number | boolean>> {
        return new Promise((resolve, reject) => {
            Excel.run(async context => {
                // Create queue of commands to get the value                

                // const cellRange = thisSheet.getRange(col + ":" + col).load("values");

                // Request the value "then" wait for it...
                context.sync().then(() => {
                    // const colValues = cellRange.values;

                    // Got the value, now resolve
                    // resolve(colValues);
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
