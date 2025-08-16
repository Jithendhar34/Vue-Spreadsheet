<template>
  <div class="app">
    <h1> Upload Spreadsheet </h1>

    <div class="controls">
      <label class="upload-btn">
        Upload File
        <input type="file" @change="handleFile" hidden />
      </label>
      <button @click="addRow">Add Row</button>
      <button @click="addColumn">Add Column</button>
      <button @click="deleteRow">Delete Last Row</button>
      <button @click="deleteColumn">Delete Last Column</button>
      <button @click="downloadFile">Download File</button>
    </div>

    <table v-if="tableData.length && tableData[0].length">
      <thead>
        <tr>
          <th v-for="(header, colIndex) in tableData[0]" :key="'header-'+colIndex">
            <input v-model="tableData[0][colIndex]" @input="saveToLocalStorage" />
          </th>
        </tr>
      </thead>
      <tbody>
        <tr v-for="(row, rowIndex) in tableData.slice(1)" :key="'row-'+rowIndex">
          <td v-for="(cell, colIndex) in row" :key="'cell-'+rowIndex+'-'+colIndex">
            <input v-model="tableData[rowIndex+1][colIndex]" @input="saveToLocalStorage" />
          </td>
        </tr>
      </tbody>
    </table>

    <p v-if="tableData.length">
      Total Rows: {{ tableData.length - 1 }}  
      Total Columns: {{ tableData[0].length }}  
      Rows Added: {{ rowsAdded }}  
      Columns Added: {{ colsAdded }}
    </p>
  </div>
</template>

<script>
import * as XLSX from "xlsx";

export default {
  name: "SpreadsheetEditor",
  data() {
    return {
      tableData: [[]],
      rowsAdded: 0,
      colsAdded: 0
    };
  },
  mounted() {
    const saved = localStorage.getItem("spreadsheetData");
    if (saved) {
      this.tableData = JSON.parse(saved);
    } else {
      this.tableData = [[""]];
    }
  },
  methods: {
    saveToLocalStorage() {
      localStorage.setItem("spreadsheetData", JSON.stringify(this.tableData));
    },
    addRow() {
      let newRow = Array(this.tableData[0].length || 1).fill("");
      this.tableData.push(newRow);
      this.rowsAdded++;
      this.saveToLocalStorage();
    },
    addColumn() {
      if (!this.tableData.length) {
        this.tableData = [[""]];
      }
      this.tableData.forEach(row => row.push(""));
      this.colsAdded++;
      this.saveToLocalStorage();
    },
    deleteRow() {
      if (this.tableData.length > 1) {
        this.tableData.pop();
        this.saveToLocalStorage();
      }
    },
    deleteColumn() {
      if (this.tableData[0].length > 0) {
        this.tableData.forEach(row => row.pop());
        this.saveToLocalStorage();
      }
    },
    handleFile(e) {
      let file = e.target.files[0];
      let reader = new FileReader();
      reader.onload = (event) => {
        let data = new Uint8Array(event.target.result);
        let workbook = XLSX.read(data, { type: 'array' });
        let sheet = workbook.Sheets[workbook.SheetNames[0]];
        this.tableData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        this.saveToLocalStorage();
      };
      reader.readAsArrayBuffer(file);
    },
    downloadFile() {
      let ws = XLSX.utils.aoa_to_sheet(this.tableData);
      let wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
      XLSX.writeFile(wb, "updated_table.xlsx");
    }
  }
};
// rglasjdf
// comment 2
</script>


<style scoped>
.app {
  background-color: black;
  color: white;
  font-family: Arial, sans-serif;
  text-align: center;
  min-height: 100vh;
}
.controls {
  margin-bottom: 15px;
}
table {
  margin: 20px auto;
  border-collapse: collapse;
}
th, td {
  border: 1px solid purple;
  padding: 8px 12px;
  min-width: 100px;
}
th {
  background-color: purple;
  color: white;
}
td {
  background-color: rgb(60, 0, 60);
}
button, .upload-btn {
  margin: 5px;
  padding: 8px 12px;
  background-color: purple;
  color: white;
  border: none;
  border-radius: 8px;
  cursor: pointer;
  font-size: 14px;
}
button:hover, .upload-btn:hover {
  background-color: rgb(100, 0, 100);
}
input {
  width: 100%;
  background: transparent;
  border: none;
  color: white;
  text-align: center;
}
input:focus {
  outline: none;
  background-color: rgba(255,255,255,0.1);
}
</style>
