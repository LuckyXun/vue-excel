<!--
 * @Author: XunL
 * @Date: 2021-07-07 22:36:41
 * @LastEditTime: 2021-07-08 00:14:06
 * @Description: file content
-->
<template>
  <div id="app">
    <input
      class="form-control fileUpload"
      type="file"
      @change="fileSelect($event)"
      name="fileUpload"
    />
  </div>
</template>

<script>
import XLSX from "xlsx";

export default {
  name: "App",
  methods: {
    fileSelect(event) {
      const [file] = event.target.files;
      const filename = file.name;
      var reader = new FileReader();
      reader.onload = function (e) {
        const fileData = e.target.result;

        const workbook = XLSX.read(fileData, { type: "binary" });
        console.log(filename);
        let offices = [];
        // 获取最近日期的科室名称
        const lastestSheet = workbook.Sheets[workbook.SheetNames[0]];
        const latestSheetJson = XLSX.utils.sheet_to_json(lastestSheet, { header: 1 });
        for (let i = 1; i < latestSheetJson.length - 1; i++) {
          offices.push(latestSheetJson[i][0]);
        }
        let count = 0;
        for (const office of offices) {
          const new_workbook = XLSX.utils.book_new();
          for (const sheetName of workbook.SheetNames) {
            const workSheet = workbook.Sheets[sheetName];
            const sheetJson = XLSX.utils.sheet_to_json(workSheet, {
              header: 1,
              raw: false,
            });
            const data = [sheetJson[0]];
            const targetLine = sheetJson.find((n) => {
              return n[0] === office;
            });
            data.push(targetLine);
            const targetsheet = XLSX.utils.aoa_to_sheet(data);
            targetsheet["!cols"] = [
              { wch: 30 },
              { wch: 15 },
              { wch: 15 },
              { wch: 15 },
              { wch: 15 },
              { wch: 15 },
              { wch: 15 },
              { wch: 15 },
              { wch: 15 },
              { wch: 15 },
              { wch: 15 },
            ];

            XLSX.utils.book_append_sheet(new_workbook, targetsheet, sheetName);
          }
          setTimeout(() => {
            XLSX.writeFile(new_workbook, `（${office}）${filename}`);
          }, count++ * 1000);
        }
      };
      reader.readAsBinaryString(file);
      event.target.value = "";
    },
  },
};
</script>

<style>
#app {
  font-family: Avenir, Helvetica, Arial, sans-serif;
  -webkit-font-smoothing: antialiased;
  -moz-osx-font-smoothing: grayscale;
  text-align: center;
  color: #2c3e50;
  margin-top: 60px;
}
</style>
