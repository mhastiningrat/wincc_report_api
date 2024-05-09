const ExcelJS = require("exceljs");
const fs = require("fs");
const moment = require("moment");
const {
	fontBolder,
	fontBold,
	textLeft,
	textCenter,
	leafHeader,
	borderThin,
} = require("../utils/excel");

const exportExcel = async (req, res) => {
	try {
		const wb = new ExcelJS.Workbook();
		const ws = wb.addWorksheet("New Sheet", {
			properties: { tabColor: { argb: "FFC0000" } },
		});

		ws.mergeCells("A1:B1");
		ws.getCell("A1").value = "LOWRATE";
		ws.getCell("A1").font = fontBolder;
		ws.getCell("A1").alignment = textLeft;

		ws.getCell("C1").value = "Date :";
		ws.getCell("C1").font = fontBold;
		ws.getCell("C1").alignment = textLeft;

		ws.mergeCells("D1:E1");
		ws.getCell("D1").value = moment().format("YYYY/MM/DD HH:mm:ss");
		ws.getCell("D1").alignment = textLeft;

		ws.getCell("F1").value = "Shift :";
		ws.getCell("F1").font = fontBold;
		ws.getCell("F1").alignment = textLeft;

		ws.mergeCells("G1:H1");
		ws.getCell("G1").value = "Pak Anas";
		ws.getCell("G1").alignment = textLeft;

		ws.getCell("A3").value = "Stream";
		ws.getCell("A3").alignment = textCenter;
		ws.getCell("A3").fill = leafHeader;
		ws.getCell("A3").font = fontBold;
		ws.getCell("A3").border = borderThin;

		ws.getCell("B3").value = "Start";
		ws.getCell("B3").alignment = textCenter;
		ws.getCell("B3").fill = leafHeader;
		ws.getCell("B3").font = fontBold;
		ws.getCell("B3").border = borderThin;

		ws.getCell("C3").value = "Stop";
		ws.getCell("C3").alignment = textCenter;
		ws.getCell("C3").fill = leafHeader;
		ws.getCell("C3").font = fontBold;
		ws.getCell("C3").border = borderThin;

		ws.getCell("D3").value = "Delay Time";
		ws.getCell("D3").alignment = textCenter;
		ws.getCell("D3").fill = leafHeader;
		ws.getCell("D3").font = fontBold;
		ws.getCell("D3").border = borderThin;

		ws.getCell("E3").value = "Category";
		ws.getCell("E3").alignment = textCenter;
		ws.getCell("E3").fill = leafHeader;
		ws.getCell("E3").font = fontBold;
		ws.getCell("E3").border = borderThin;

		ws.getCell("F3").value = "Ent Name";
		ws.getCell("F3").alignment = textCenter;
		ws.getCell("F3").fill = leafHeader;
		ws.getCell("F3").font = fontBold;
		ws.getCell("F3").border = borderThin;

		ws.getCell("G3").value = "D.Code";
		ws.getCell("G3").alignment = textCenter;
		ws.getCell("G3").fill = leafHeader;
		ws.getCell("G3").font = fontBold;
		ws.getCell("G3").border = borderThin;

		ws.getCell("H3").value = "Comments";
		ws.getCell("H3").alignment = textCenter;
		ws.getCell("H3").fill = leafHeader;
		ws.getCell("H3").font = fontBold;
		ws.getCell("H3").border = borderThin;

		let data = [
			{
				stream: "Stream 001",
				start: moment().format("YYYY/MM/DD HH:mm:ss"),
				stop: moment().format("YYYY/MM/DD HH:mm:ss"),
				delay_time: "1 jam",
				category: "medium",
				ent_name: "",
				delay_code: "unknown code",
				comment: "masih dalam perbaikan, due date minggu depan",
			},
			{
				stream: "Stream 001",
				start: moment().format("YYYY/MM/DD HH:mm:ss"),
				stop: moment().format("YYYY/MM/DD HH:mm:ss"),
				delay_time: "1 jam",
				category: "medium",
				ent_name: "",
				delay_code: "unknown code",
				comment: "masih dalam perbaikan, due date minggu depan",
			},
			{
				stream: "Stream 001",
				start: moment().format("YYYY/MM/DD HH:mm:ss"),
				stop: moment().format("YYYY/MM/DD HH:mm:ss"),
				delay_time: "1 jam",
				category: "medium",
				ent_name: "",
				delay_code: "unknown code",
				comment: "masih dalam perbaikan, due date minggu depan",
			},
			{
				stream: "Stream 001",
				start: moment().format("YYYY/MM/DD HH:mm:ss"),
				stop: moment().format("YYYY/MM/DD HH:mm:ss"),
				delay_time: "1 jam",
				category: "medium",
				ent_name: "",
				delay_code: "unknown code",
				comment: "masih dalam perbaikan, due date minggu depan",
			},
			{
				stream: "Stream 001",
				start: moment().format("YYYY/MM/DD HH:mm:ss"),
				stop: moment().format("YYYY/MM/DD HH:mm:ss"),
				delay_time: "1 jam",
				category: "medium",
				ent_name: "",
				delay_code: "unknown code",
				comment: "masih dalam perbaikan, due date minggu depan",
			},
			{
				stream: "Stream 001",
				start: moment().format("YYYY/MM/DD HH:mm:ss"),
				stop: moment().format("YYYY/MM/DD HH:mm:ss"),
				delay_time: "1 jam",
				category: "medium",
				ent_name: "",
				delay_code: "unknown code",
				comment: "masih dalam perbaikan, due date minggu depan",
			},
		];

		let startData = 3;
		for (let i in data) {
			startData++;
			ws.getCell("A" + startData).value = data[i].stream;
			ws.getCell("A" + startData).alignment = textCenter;
			ws.getCell("A" + startData).border = borderThin;

			ws.getCell("B" + startData).value = data[i].start;
			ws.getCell("B" + startData).alignment = textCenter;
			ws.getCell("B" + startData).border = borderThin;

			ws.getCell("C" + startData).value = data[i].stop;
			ws.getCell("C" + startData).alignment = textCenter;
			ws.getCell("C" + startData).border = borderThin;

			ws.getCell("D" + startData).value = data[i].delay_time;
			ws.getCell("D" + startData).alignment = textCenter;
			ws.getCell("D" + startData).border = borderThin;

			ws.getCell("E" + startData).value = data[i].category;
			ws.getCell("E" + startData).alignment = textCenter;
			ws.getCell("E" + startData).border = borderThin;

			ws.getCell("F" + startData).value = data[i].ent_name;
			ws.getCell("F" + startData).alignment = textCenter;
			ws.getCell("F" + startData).border = borderThin;

			ws.getCell("G" + startData).value = data[i].delay_code;
			ws.getCell("G" + startData).alignment = textCenter;
			ws.getCell("G" + startData).border = borderThin;

			ws.getCell("H" + startData).value = data[i].comment;
			ws.getCell("H" + startData).alignment = textCenter;
			ws.getCell("H" + startData).border = borderThin;
		}

		let datePrint = `${moment().format("YYYYMMDD")}_${moment().format(
			"HHmmss"
		)}`;
		const direct = "/tmp";
		await wb.xlsx
			.writeFile(`${direct}/LOWRATE_REPORT_${datePrint}.xlsx`)
			.then(() => {
				res.download(
					`${direct}/LOWRATE_REPORT_${datePrint}.xlsx`,
					`LOWRATE_REPORT_${datePrint}.xlsx`,
					(err) => {
						if (err) {
							console.log(err);
						} else {
							fs.unlinkSync(`${direct}/LOWRATE_REPORT_${datePrint}.xlsx`);
						}
					}
				);
			});
	} catch (error) {
		res.json({ status: "failed", reason: error });
	}
};

module.exports = {
	exportExcel,
};
