import "./styles.css";
import * as ExcelJs from "exceljs";
import * as XLSX from "xlsx/xlsx.mjs";
import { saveAs } from "file-saver";
import * as FileSaver from "file-saver";
export default function App() {
  async function exportExcel() {
    const workbook = new ExcelJs.Workbook();
    const worksheet = workbook.addWorksheet("My Sheet");

    //Cách 1 ------------------------
    //tiêu đề cột
    // worksheet.columns = [
    //   { header: "Id", key: "id", width: 10 },
    //   { header: "Name", key: "name", width: 32 },
    //   { header: "D.O.B.", key: "dob", width: 15 }
    // ];

    // //các hàng được add vào sheet
    // worksheet.addRow({ id: 1, name: "John Doe", dob: new Date(1970, 1, 1) });
    // worksheet.addRow({ id: 2, name: "Jane Doe", dob: new Date(1965, 1, 7) });
    //Cách 1 ------------------------

    //Cách 2 ------------------------
    worksheet.addTable({
      name: "MyTable",
      ref: "A1",
      // headerRow: true,
      totalsRow: true,
      style: {
        // theme: 'TableStyleLight9',
        showRowStripes: true
      },
      columns: [
        { name: "Date", totalsRowLabel: "Totals:", filterButton: true },
        { name: "Amount", totalsRowFunction: "sum", filterButton: false }
      ],
      rows: [
        [new Date("2019-07-20"), 70.1],
        [new Date("2019-07-21"), 70.6],
        [new Date("2019-07-22"), 70.1]
      ]
    });
    //Cách 2 ------------------------

    //border cell
    for (let i = 1; i <= 3; i++) {
      worksheet.getCell(`A${i}`).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" }
      };
    }

    // save under export.xlsx
    workbook.xlsx
      .writeBuffer()
      .then((buffer) =>
        FileSaver.saveAs(new Blob([buffer]), `${Date.now()}_feedback.xlsx`)
      )
      .catch((err) => console.log("Error writing excel export", err));
  }

  return (
    <div className="App">
      <button onClick={exportExcel}>Tải file</button>
    </div>
  );
}
