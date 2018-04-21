export class DataAccess {

    public getSheet(name: string): Promise<Excel.Worksheet> {
        // Function Returns a promise of a value in the future
        return new Promise( (resolve, reject) => {
            // Go fetch the worksheet... this won't happen immediately...
            Excel.run( async context => {

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

    public getValue(row: number, col: number): Promise<Excel.Worksheet> {
        // Function Returns a promise of a value in the future
        return new Promise( (resolve, reject) => {
            // Go fetch the worksheet... this won't happen immediately...
            Excel.run( async context => {

                // Create queue of commands to get the value                

                // Request the value "then" wait for it...
                context.sync().then(() => {
                    // Got the value, now resolve
                });
            });
        });
    }

    public emptyPromise(): Promise<string> {
        return new Promise( (resolve, reject) => {
            resolve('asdf');
        });
    }
}
