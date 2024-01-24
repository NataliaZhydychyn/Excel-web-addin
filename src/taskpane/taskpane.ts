import 'reflect-metadata';
import { injectable, Container } from 'inversify';

interface IExcelService {
    insertText(inputValue: string): Promise<void>;
    handleSelectionChanged(): Promise<void>;
    clearCell(context: Excel.RequestContext, address: string): void;
}

@injectable()
class ExcelService implements IExcelService {
    private previousCellAddress: string | undefined;

    constructor() {}

    async insertText(inputValue: string): Promise<void> {
        try {
            await Excel.run(async (context) => {
                document.getElementById("textInput") as HTMLInputElement;
                const selectedCell = context.workbook.getActiveCell();

                if (selectedCell) {
                    selectedCell.values = [[inputValue]];
                }
                await context.sync();
            });
        } catch (error) {
            console.error('Error inserting text:', error);
        }
    }

    async handleSelectionChanged(): Promise<void> {
        try {
            const textInput = document.querySelector("#textInput") as HTMLInputElement;
            textInput.value = "";

            await Excel.run(async (context) => {
                const currentCell = context.workbook.getSelectedRange();
                currentCell.load("address");
                await context.sync();

                if (this.previousCellAddress) {
                    this.clearCell(context, this.previousCellAddress);
                    await context.sync();
                }

                this.previousCellAddress = currentCell.address;
            });
        } catch (error) {
            console.error('Error handling selection change:', error);
        }
    }

    clearCell(context: Excel.RequestContext, address: string): void {
        const activeSheet = context.workbook.worksheets.getActiveWorksheet();
        const cell = activeSheet.getRange(address);
        cell.clear();
    }
}

// Configure the InversifyJS container
const container = new Container();
container.bind<IExcelService>(ExcelService).to(ExcelService);

Office.onReady(async (info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("sideload-msg")!.style.display = "none";
        document.getElementById("app-body")!.style.display = "flex";

        const excelService = container.get<IExcelService>(ExcelService);

        const textInput = document.querySelector("#textInput") as HTMLInputElement;
        textInput.addEventListener("input", async () => {
            await excelService.insertText(textInput.value);
        });

        await Excel.run(async (context) => {
            context.workbook.onSelectionChanged.add(() => excelService.handleSelectionChanged());
            await context.sync();
        });
    }
});
