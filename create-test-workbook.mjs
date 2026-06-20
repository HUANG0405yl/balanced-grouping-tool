import * as XLSX from "xlsx";
import * as fs from "node:fs/promises";

const headers = ["姓名", "公司", "性别", "年龄", "部门", "岗位"];

const companies = ["华东公司", "华南公司", "西北公司"];
const genders = ["男", "女"];

const rows = Array.from({ length: 23 }, (_, i) => {
  return [
    `员工${String(i + 1).padStart(2, "0")}`,
    companies[i % companies.length],
    genders[i % 2],
    22 + (i % 10),
    ["财务部", "市场部", "运营部", "技术部"][i % 4],
    ["经理", "主管", "专员"][i % 3]
  ];
});

const data = [headers, ...rows];

const ws = XLSX.utils.aoa_to_sheet(data);
const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, ws, "人员名单");

await fs.mkdir("./work", { recursive: true });

const out = "./work/test-people.xlsx";

const buffer = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
await fs.writeFile(out, buffer);

console.log("生成成功：", out);

