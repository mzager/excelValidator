
export class DataAccess {

    // private _excel: Excel.RequestContext;

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

    public emptyPromise(): Promise<string> {
        return new Promise( (resolve, reject) => {
            resolve('asdf');
        });
    }
    constructor(excel: Excel.RequestContext) {
        // this._excel = excel;
    }
}
