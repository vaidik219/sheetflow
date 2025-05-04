import { google } from "googleapis";
import { existsSync } from "fs";

interface SheetConfig {
  name: string;
  id: string;
  range: string;
}

export default class SheetFlow {
  sheetsAPI: any;

  constructor(private credentialsPath: string, private sheet: SheetConfig) {
    if (!existsSync(credentialsPath))
      throw new Error(`Credentials cannot be found at ${credentialsPath}.`);

    this.sheetsAPI = google.sheets({
      version: "v4",
      auth: new google.auth.GoogleAuth({
        keyFile: this.credentialsPath,
        scopes: ["https://www.googleapis.com/auth/spreadsheets"],
      }),
    });
    console.log("Authorized.");
  }

  async read() {
    const spreadsheetId = this.sheet.id;
    const range = `${this.sheet.name}!${this.sheet.range}`;
    const response = await this.sheetsAPI.spreadsheets.values.get({
      spreadsheetId,
      range,
    });
    const rows = response.data.values;
    console.log(rows);

    if (!rows) return "No data found.";

    const [headers, ...dataRows] = rows;
    return dataRows.map((row: string[]) => {
      const obj = {};
      row.forEach((value: string, index: number) => {
        const key = headers[index];
        (obj as Record<string, string>)[key] = value;
      });
      return obj;
    });
  }

  async search(query: Record<string, any>): Promise<Record<string, any>[]> {
    const values = await this.read();

    const matches = values.filter((obj: any) => {
      return Object.entries(query).every(([key, value]) => {
        return obj.hasOwnProperty(key) && obj[key] === value;
      });
    });

    return matches.length ? matches : [];
  }

  async write(data: string[]) {
    const spreadsheetId = this.sheet.id;
    const range = `${this.sheet.name}!${this.sheet.range}`;
    this.sheetsAPI.spreadsheets.values.append({
      spreadsheetId,
      range,
      valueInputOption: "USER_ENTERED",
      resource: {
        values: [data],
      },
    });
    return "Wrote.";
  }

  async update(data: Record<string, any>) {
    const spreadsheetId = this.sheet.id;
    const range = `${this.sheet.name}!${this.sheet.range}`;
    try {
      const response = await this.sheetsAPI.spreadsheets.values.get({
        spreadsheetId,
        range: range,
      });

      const rows = response.data.values;

      if (rows.length === 0) return "No data found.";

      const headers = rows[0];
      const columnIndex = headers.indexOf(Object.keys(data)[0]);

      if (columnIndex === -1) return "Column not found.";

      for (let i = 1; i < rows.length; i++) {
        if (rows[i][columnIndex] === data[Object.keys(data)[0]]) {
          for (let j = 1; j < headers.length; j++) {
            if (Object.keys(data).includes(headers[j])) {
              rows[i][j] = data[headers[j]];
            }
          }
          break;
        }
      }

      this.sheetsAPI.spreadsheets.values.update({
        spreadsheetId,
        range: range,
        valueInputOption: "RAW",
        resource: {
          values: rows,
        },
      });

      return "Updated.";
    } catch (e) {
      console.error("Error updating data: ", e);
    }
  }

  async delete(
    data: Record<string, any>,
    all: boolean = false
  ): Promise<string> {
    const spreadsheetId = this.sheet.id;
    const range = `${this.sheet.name}!${this.sheet.range}`;

    try {
      const spreadsheet = await this.sheetsAPI.spreadsheets.get({
        spreadsheetId,
        ranges: [range],
        fields: "sheets(properties(sheetId,title))",
      });
      const sheetId = spreadsheet.data.sheets[0].properties.sheetId;

      const response = await this.sheetsAPI.spreadsheets.values.get({
        spreadsheetId,
        range,
      });

      const rows: string[][] = response.data.values;
      if (!rows || rows.length === 0) return "No data found.";

      const headers: string[] = rows[0];
      const columnIndex: number = headers.indexOf(Object.keys(data)[0]);
      if (columnIndex === -1) return "Column not found.";

      const rowIndices: number[] = [];
      for (let i = rows.length - 1; i >= 1; i--) {
        if (rows[i][columnIndex] === data[Object.keys(data)[0]]) {
          rowIndices.push(i);
          if (!all) break;
        }
      }

      if (rowIndices.length === 0) return "No matching data found.";

      await this.sheetsAPI.spreadsheets.batchUpdate({
        spreadsheetId,
        resource: {
          requests: rowIndices.map((index) => ({
            deleteDimension: {
              range: {
                sheetId,
                dimension: "ROWS",
                startIndex: index,
                endIndex: index + 1,
              },
            },
          })),
        },
      });

      return `Deleted ${rowIndices.length} row(s).`;
    } catch (e) {
      console.error("Error deleting data: ", e);
      throw new Error("Failed to delete data");
    }
  }
}
