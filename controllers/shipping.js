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
	oceanHeader,
	borderBold,
	oceanMudaHeader,
	toscaHeader,
} = require("../utils/excel");

const exportExcel = async (req, res) => {
	try {
		const wb = new ExcelJS.Workbook();
		const ws = wb.addWorksheet("New Sheet", {
			properties: { tabColor: { argb: "FFC0000" } },
		});

		ws.mergeCells("A1:I1");
		ws.getCell("A1").value = "KALTIM PRIMA COAL";
		ws.getCell("A1").font = fontBolder;
		ws.getCell("A1").alignment = textCenter;
		ws.getCell("A1").fill = oceanHeader;
		ws.getCell("A1").border = borderBold;

		ws.mergeCells("A2:I2");
		ws.getCell("A2").value = "SHIPPING REPORT - VESSEL LOADING";
		ws.getCell("A2").font = fontBolder;
		ws.getCell("A2").alignment = textCenter;
		ws.getCell("A2").fill = oceanHeader;
		ws.getCell("A2").border = borderBold;

		ws.getCell("A4").value = "SHIP NAME";
		ws.getCell("A4").alignment = textLeft;
		ws.getCell("A4").fill = oceanMudaHeader;

		ws.getCell("B4").value = ":";
		ws.getCell("B4").alignment = textCenter;
		ws.getCell("B4").fill = oceanMudaHeader;

		ws.mergeCells("C4:I4");
		ws.getCell("C4").alignment = textLeft;
		ws.getCell("C4").fill = oceanMudaHeader;

		ws.getCell("A5").value = "NO OF HATCH";
		ws.getCell("A5").alignment = textLeft;
		ws.getCell("A5").fill = toscaHeader;

		ws.getCell("B5").value = ":";
		ws.getCell("B5").alignment = textCenter;
		ws.getCell("B5").fill = toscaHeader;

		ws.mergeCells("C5:I5");
		ws.getCell("C5").alignment = textLeft;
		ws.getCell("C5").fill = toscaHeader;

		ws.getCell("A6").value = "SHIP CAPACITY";
		ws.getCell("A6").alignment = textLeft;
		ws.getCell("A6").fill = oceanMudaHeader;

		ws.getCell("B6").value = ":";
		ws.getCell("B6").alignment = textCenter;
		ws.getCell("B6").fill = oceanMudaHeader;

		ws.mergeCells("C6:I6");
		ws.getCell("C6").alignment = textLeft;
		ws.getCell("C6").fill = oceanMudaHeader;

		ws.getCell("A7").value = "TOTAL TONNES LOADED";
		ws.getCell("A7").alignment = textLeft;
		ws.getCell("A7").fill = toscaHeader;

		ws.getCell("B7").value = ":";
		ws.getCell("B7").alignment = textCenter;
		ws.getCell("B7").fill = toscaHeader;

		ws.mergeCells("C7:I7");
		ws.getCell("C7").alignment = textLeft;
		ws.getCell("C7").fill = toscaHeader;

		ws.getCell("A9").value = "SHIP BERTHED DATE";
		ws.getCell("A9").alignment = textLeft;
		ws.getCell("A9").fill = oceanMudaHeader;

		ws.getCell("B9").value = ":";
		ws.getCell("B9").alignment = textCenter;
		ws.getCell("B9").fill = oceanMudaHeader;

		ws.mergeCells("C9:D9");
		ws.getCell("C9").alignment = textLeft;
		ws.getCell("C9").fill = oceanMudaHeader;

		ws.getCell("F9").value = "SHIP BERTHED TIME";
		ws.getCell("F9").alignment = textLeft;
		ws.getCell("F9").fill = oceanMudaHeader;

		ws.getCell("G9").value = ":";
		ws.getCell("G9").alignment = textCenter;
		ws.getCell("G9").fill = oceanMudaHeader;

		ws.mergeCells("H9:I9");
		ws.getCell("H9").alignment = textLeft;
		ws.getCell("H9").fill = oceanMudaHeader;

		ws.getCell("A10").value = "LOADING START DATE";
		ws.getCell("A10").alignment = textLeft;
		ws.getCell("A10").fill = toscaHeader;

		ws.getCell("B10").value = ":";
		ws.getCell("B10").alignment = textCenter;
		ws.getCell("B10").fill = toscaHeader;

		ws.mergeCells("C10:D10");
		ws.getCell("C10").alignment = textLeft;
		ws.getCell("C10").fill = toscaHeader;

		ws.getCell("F10").value = "LOADING START TIME";
		ws.getCell("F10").alignment = textLeft;
		ws.getCell("F10").fill = toscaHeader;

		ws.getCell("G10").value = ":";
		ws.getCell("G10").alignment = textCenter;
		ws.getCell("G10").fill = toscaHeader;

		ws.mergeCells("H10:I10");
		ws.getCell("H10").alignment = textLeft;
		ws.getCell("H10").fill = toscaHeader;

		ws.getCell("A12").value = "";
		ws.getCell("A12").alignment = textLeft;
		ws.getCell("A12").fill = oceanMudaHeader;

		ws.getCell("B12").value = "";
		ws.getCell("B12").alignment = textCenter;
		ws.getCell("B12").fill = oceanMudaHeader;

		ws.mergeCells("C12:D12");
		ws.getCell("C12").alignment = textLeft;
		ws.getCell("C12").fill = oceanMudaHeader;

		ws.getCell("F12").value = "";
		ws.getCell("F12").alignment = textLeft;
		ws.getCell("F12").fill = oceanMudaHeader;

		ws.getCell("G12").value = "";
		ws.getCell("G12").alignment = textCenter;
		ws.getCell("G12").fill = oceanMudaHeader;

		ws.mergeCells("H12:I12");
		ws.getCell("H12").alignment = textLeft;
		ws.getCell("H12").fill = oceanMudaHeader;

		ws.getCell("A13").value = "";
		ws.getCell("A13").alignment = textLeft;
		ws.getCell("A13").fill = toscaHeader;

		ws.getCell("B13").value = "";
		ws.getCell("B13").alignment = textCenter;
		ws.getCell("B13").fill = toscaHeader;

		ws.mergeCells("C13:D13");
		ws.getCell("C13").alignment = textLeft;
		ws.getCell("C13").fill = toscaHeader;

		ws.getCell("F13").value = "";
		ws.getCell("F13").alignment = textLeft;
		ws.getCell("F13").fill = toscaHeader;

		ws.getCell("G13").value = "";
		ws.getCell("G13").alignment = textCenter;
		ws.getCell("G13").fill = toscaHeader;

		ws.mergeCells("H13:I13");
		ws.getCell("H13").alignment = textLeft;
		ws.getCell("H13").fill = toscaHeader;

		ws.getCell("A14").value = "";
		ws.getCell("A14").alignment = textLeft;
		ws.getCell("A14").fill = oceanMudaHeader;

		ws.getCell("B14").value = "";
		ws.getCell("B14").alignment = textCenter;
		ws.getCell("B14").fill = oceanMudaHeader;

		ws.mergeCells("C14:D14");
		ws.getCell("C14").alignment = textLeft;
		ws.getCell("C14").fill = oceanMudaHeader;

		ws.getCell("F14").value = "";
		ws.getCell("F14").alignment = textLeft;
		ws.getCell("F14").fill = oceanMudaHeader;

		ws.getCell("G14").value = "";
		ws.getCell("G14").alignment = textCenter;
		ws.getCell("G14").fill = oceanMudaHeader;

		ws.mergeCells("H14:I14");
		ws.getCell("H14").alignment = textLeft;
		ws.getCell("H14").fill = oceanMudaHeader;

		ws.getCell("A15").value = "";
		ws.getCell("A15").alignment = textLeft;
		ws.getCell("A15").fill = toscaHeader;

		ws.getCell("B15").value = "";
		ws.getCell("B15").alignment = textCenter;
		ws.getCell("B15").fill = toscaHeader;

		ws.mergeCells("C15:D15");
		ws.getCell("C15").alignment = textLeft;
		ws.getCell("C15").fill = toscaHeader;

		ws.getCell("F15").value = "";
		ws.getCell("F15").alignment = textLeft;
		ws.getCell("F15").fill = toscaHeader;

		ws.getCell("G15").value = "";
		ws.getCell("G15").alignment = textCenter;
		ws.getCell("G15").fill = toscaHeader;

		ws.mergeCells("H15:I15");
		ws.getCell("H15").alignment = textLeft;
		ws.getCell("H15").fill = toscaHeader;

		ws.getCell("A16").value = "";
		ws.getCell("A16").alignment = textLeft;
		ws.getCell("A16").fill = oceanMudaHeader;

		ws.getCell("B16").value = "";
		ws.getCell("B16").alignment = textCenter;
		ws.getCell("B16").fill = oceanMudaHeader;

		ws.mergeCells("C16:D16");
		ws.getCell("C16").alignment = textLeft;
		ws.getCell("C16").fill = oceanMudaHeader;

		ws.getCell("F16").value = "";
		ws.getCell("F16").alignment = textLeft;
		ws.getCell("F16").fill = oceanMudaHeader;

		ws.getCell("G16").value = "";
		ws.getCell("G16").alignment = textCenter;
		ws.getCell("G16").fill = oceanMudaHeader;

		ws.mergeCells("H16:I16");
		ws.getCell("H16").alignment = textLeft;
		ws.getCell("H16").fill = oceanMudaHeader;

		ws.getCell("A17").value = "";
		ws.getCell("A17").alignment = textLeft;
		ws.getCell("A17").fill = toscaHeader;

		ws.getCell("B17").value = "";
		ws.getCell("B17").alignment = textCenter;
		ws.getCell("B17").fill = toscaHeader;

		ws.mergeCells("C17:D17");
		ws.getCell("C17").alignment = textLeft;
		ws.getCell("C17").fill = toscaHeader;

		ws.getCell("F17").value = "";
		ws.getCell("F17").alignment = textLeft;
		ws.getCell("F17").fill = toscaHeader;

		ws.getCell("G17").value = "";
		ws.getCell("G17").alignment = textCenter;
		ws.getCell("G17").fill = toscaHeader;

		ws.mergeCells("H17:I17");
		ws.getCell("H17").alignment = textLeft;
		ws.getCell("H17").fill = toscaHeader;

		ws.getCell("A18").value = "";
		ws.getCell("A18").alignment = textLeft;
		ws.getCell("A18").fill = oceanMudaHeader;

		ws.getCell("B18").value = "";
		ws.getCell("B18").alignment = textCenter;
		ws.getCell("B18").fill = oceanMudaHeader;

		ws.mergeCells("C18:D18");
		ws.getCell("C18").alignment = textLeft;
		ws.getCell("C18").fill = oceanMudaHeader;

		ws.getCell("F18").value = "";
		ws.getCell("F18").alignment = textLeft;
		ws.getCell("F18").fill = oceanMudaHeader;

		ws.getCell("G18").value = "";
		ws.getCell("G18").alignment = textCenter;
		ws.getCell("G18").fill = oceanMudaHeader;

		ws.mergeCells("H18:I18");
		ws.getCell("H18").alignment = textLeft;
		ws.getCell("H18").fill = oceanMudaHeader;

		ws.getCell("A19").value = "";
		ws.getCell("A19").alignment = textLeft;
		ws.getCell("A19").fill = toscaHeader;

		ws.getCell("B19").value = "";
		ws.getCell("B19").alignment = textCenter;
		ws.getCell("B19").fill = toscaHeader;

		ws.mergeCells("C19:D19");
		ws.getCell("C19").alignment = textLeft;
		ws.getCell("C19").fill = toscaHeader;

		ws.getCell("F19").value = "";
		ws.getCell("F19").alignment = textLeft;
		ws.getCell("F19").fill = toscaHeader;

		ws.getCell("G19").value = "";
		ws.getCell("G19").alignment = textCenter;
		ws.getCell("G19").fill = toscaHeader;

		ws.mergeCells("H19:I19");
		ws.getCell("H19").alignment = textLeft;
		ws.getCell("H19").fill = toscaHeader;

		ws.getCell("A20").value = "";
		ws.getCell("A20").alignment = textLeft;
		ws.getCell("A20").fill = oceanMudaHeader;

		ws.getCell("B20").value = "";
		ws.getCell("B20").alignment = textCenter;
		ws.getCell("B20").fill = oceanMudaHeader;

		ws.mergeCells("C20:D20");
		ws.getCell("C20").alignment = textLeft;
		ws.getCell("C20").fill = oceanMudaHeader;

		ws.getCell("F20").value = "";
		ws.getCell("F20").alignment = textLeft;
		ws.getCell("F20").fill = oceanMudaHeader;

		ws.getCell("G20").value = "";
		ws.getCell("G20").alignment = textCenter;
		ws.getCell("G20").fill = oceanMudaHeader;

		ws.mergeCells("H20:I20");
		ws.getCell("H20").alignment = textLeft;
		ws.getCell("H20").fill = oceanMudaHeader;
		//================================================================
		ws.getCell("A21").value = "";
		ws.getCell("A21").alignment = textLeft;
		ws.getCell("A21").fill = toscaHeader;

		ws.getCell("B21").value = "";
		ws.getCell("B21").alignment = textCenter;
		ws.getCell("B21").fill = toscaHeader;

		ws.mergeCells("C21:D21");
		ws.getCell("C21").alignment = textLeft;
		ws.getCell("C21").fill = toscaHeader;

		ws.getCell("F21").value = "";
		ws.getCell("F21").alignment = textLeft;
		ws.getCell("F21").fill = toscaHeader;

		ws.getCell("G21").value = "";
		ws.getCell("G21").alignment = textCenter;
		ws.getCell("G21").fill = toscaHeader;

		ws.mergeCells("H21:I21");
		ws.getCell("H21").alignment = textLeft;
		ws.getCell("H21").fill = toscaHeader;

		ws.getCell("A22").value = "";
		ws.getCell("A22").alignment = textLeft;
		ws.getCell("A22").fill = oceanMudaHeader;

		ws.getCell("B22").value = "";
		ws.getCell("B22").alignment = textCenter;
		ws.getCell("B22").fill = oceanMudaHeader;

		ws.mergeCells("C22:D22");
		ws.getCell("C22").alignment = textLeft;
		ws.getCell("C22").fill = oceanMudaHeader;

		ws.getCell("F22").value = "";
		ws.getCell("F22").alignment = textLeft;
		ws.getCell("F22").fill = oceanMudaHeader;

		ws.getCell("G22").value = "";
		ws.getCell("G22").alignment = textCenter;
		ws.getCell("G22").fill = oceanMudaHeader;

		ws.mergeCells("H22:I22");
		ws.getCell("H22").alignment = textLeft;
		ws.getCell("H22").fill = oceanMudaHeader;

		ws.getCell("A23").value = "";
		ws.getCell("A23").alignment = textLeft;
		ws.getCell("A23").fill = toscaHeader;

		ws.getCell("B23").value = "";
		ws.getCell("B23").alignment = textCenter;
		ws.getCell("B23").fill = toscaHeader;

		ws.mergeCells("C23:D23");
		ws.getCell("C23").alignment = textLeft;
		ws.getCell("C23").fill = toscaHeader;

		ws.getCell("F23").value = "";
		ws.getCell("F23").alignment = textLeft;
		ws.getCell("F23").fill = toscaHeader;

		ws.getCell("G23").value = "";
		ws.getCell("G23").alignment = textCenter;
		ws.getCell("G23").fill = toscaHeader;

		ws.mergeCells("H23:I23");
		ws.getCell("H23").alignment = textLeft;
		ws.getCell("H23").fill = toscaHeader;

		ws.getCell("A24").value = "";
		ws.getCell("A24").alignment = textLeft;
		ws.getCell("A24").fill = oceanMudaHeader;

		ws.getCell("B24").value = "";
		ws.getCell("B24").alignment = textCenter;
		ws.getCell("B24").fill = oceanMudaHeader;

		ws.mergeCells("C24:D24");
		ws.getCell("C24").alignment = textLeft;
		ws.getCell("C24").fill = oceanMudaHeader;

		ws.getCell("F24").value = "";
		ws.getCell("F24").alignment = textLeft;
		ws.getCell("F24").fill = oceanMudaHeader;

		ws.getCell("G24").value = "";
		ws.getCell("G24").alignment = textCenter;
		ws.getCell("G24").fill = oceanMudaHeader;

		ws.mergeCells("H24:I24");
		ws.getCell("H24").alignment = textLeft;
		ws.getCell("H24").fill = oceanMudaHeader;

		ws.getCell("A26").value = "LOADING FINISHED DATE";
		ws.getCell("A26").alignment = textLeft;
		ws.getCell("A26").fill = oceanMudaHeader;

		ws.getCell("B26").value = ":";
		ws.getCell("B26").alignment = textCenter;
		ws.getCell("B26").fill = oceanMudaHeader;

		ws.mergeCells("C26:D26");
		ws.getCell("C26").alignment = textLeft;
		ws.getCell("C26").fill = oceanMudaHeader;

		ws.getCell("F26").value = "LOADING FINISHED TIME";
		ws.getCell("F26").alignment = textLeft;
		ws.getCell("F26").fill = oceanMudaHeader;

		ws.getCell("G26").value = ":";
		ws.getCell("G26").alignment = textCenter;
		ws.getCell("G26").fill = oceanMudaHeader;

		ws.mergeCells("H26:I26");
		ws.getCell("H26").alignment = textLeft;
		ws.getCell("H26").fill = oceanMudaHeader;

		ws.getCell("A27").value = "SHIP DEBERTH DATE";
		ws.getCell("A27").alignment = textLeft;
		ws.getCell("A27").fill = toscaHeader;

		ws.getCell("B27").value = ":";
		ws.getCell("B27").alignment = textCenter;
		ws.getCell("B27").fill = toscaHeader;

		ws.mergeCells("C27:D27");
		ws.getCell("C27").alignment = textLeft;
		ws.getCell("C27").fill = toscaHeader;

		ws.getCell("F27").value = "SHIP DEBERTH TIME";
		ws.getCell("F27").alignment = textLeft;
		ws.getCell("F27").fill = toscaHeader;

		ws.getCell("G27").value = ":";
		ws.getCell("G27").alignment = textCenter;
		ws.getCell("G27").fill = toscaHeader;

		ws.mergeCells("H27:I27");
		ws.getCell("H27").alignment = textLeft;
		ws.getCell("H27").fill = toscaHeader;

		ws.getCell("A28").value = "SHIP DEPARTH DATE";
		ws.getCell("A28").alignment = textLeft;
		ws.getCell("A28").fill = oceanMudaHeader;

		ws.getCell("B28").value = ":";
		ws.getCell("B28").alignment = textCenter;
		ws.getCell("B28").fill = oceanMudaHeader;

		ws.mergeCells("C28:D28");
		ws.getCell("C28").alignment = textLeft;
		ws.getCell("C28").fill = oceanMudaHeader;

		ws.getCell("F28").value = "SHIP DEPARTH TIME";
		ws.getCell("F28").alignment = textLeft;
		ws.getCell("F28").fill = oceanMudaHeader;

		ws.getCell("G28").value = ":";
		ws.getCell("G28").alignment = textCenter;
		ws.getCell("G28").fill = oceanMudaHeader;

		ws.mergeCells("H28:I28");
		ws.getCell("H28").alignment = textLeft;
		ws.getCell("H28").fill = oceanMudaHeader;

		ws.getCell("A30").value = "TRESTLE BELT SCALE";
		ws.getCell("A30").alignment = textLeft;
		ws.getCell("A30").fill = oceanMudaHeader;

		ws.getCell("B30").value = ":";
		ws.getCell("B30").alignment = textCenter;
		ws.getCell("B30").fill = oceanMudaHeader;

		ws.mergeCells("C30:D30");
		ws.getCell("C30").alignment = textLeft;
		ws.getCell("C30").fill = oceanMudaHeader;

		ws.getCell("F30").value = "TRESTLE ERROR";
		ws.getCell("F30").alignment = textLeft;
		ws.getCell("F30").fill = oceanMudaHeader;

		ws.getCell("G30").value = ":";
		ws.getCell("G30").alignment = textCenter;
		ws.getCell("G30").fill = oceanMudaHeader;

		ws.mergeCells("H30:I30");
		ws.getCell("H30").alignment = textLeft;
		ws.getCell("H30").fill = oceanMudaHeader;

		ws.getCell("A31").value = "NORTH TRANSFER BELT SCALE";
		ws.getCell("A31").alignment = textLeft;
		ws.getCell("A31").fill = toscaHeader;

		ws.getCell("B31").value = ":";
		ws.getCell("B31").alignment = textCenter;
		ws.getCell("B31").fill = toscaHeader;

		ws.mergeCells("C31:D31");
		ws.getCell("C31").alignment = textLeft;
		ws.getCell("C31").fill = toscaHeader;

		ws.getCell("F31").value = "TRANSFER ERROR";
		ws.getCell("F31").alignment = textLeft;
		ws.getCell("F31").fill = toscaHeader;

		ws.getCell("G31").value = ":";
		ws.getCell("G31").alignment = textCenter;
		ws.getCell("G31").fill = toscaHeader;

		ws.mergeCells("H31:I31");
		ws.getCell("H31").alignment = textLeft;
		ws.getCell("H31").fill = toscaHeader;

		ws.getCell("A32").value = "SOUTH TRANSFER BELT SCALE";
		ws.getCell("A32").alignment = textLeft;
		ws.getCell("A32").fill = oceanMudaHeader;

		ws.getCell("B32").value = ":";
		ws.getCell("B32").alignment = textCenter;
		ws.getCell("B32").fill = oceanMudaHeader;

		ws.mergeCells("C32:D32");
		ws.getCell("C32").alignment = textLeft;
		ws.getCell("C32").fill = oceanMudaHeader;

		ws.getCell("F32").value = "";
		ws.getCell("F32").alignment = textLeft;
		ws.getCell("F32").fill = oceanMudaHeader;

		ws.getCell("G32").value = "";
		ws.getCell("G32").alignment = textCenter;
		ws.getCell("G32").fill = oceanMudaHeader;

		ws.mergeCells("H32:I32");
		ws.getCell("H32").alignment = textLeft;
		ws.getCell("H32").fill = oceanMudaHeader;

		ws.getCell("A33").value = "DRAFT SURVEY";
		ws.getCell("A33").alignment = textLeft;
		ws.getCell("A33").fill = toscaHeader;

		ws.getCell("B33").value = ":";
		ws.getCell("B33").alignment = textCenter;
		ws.getCell("B33").fill = toscaHeader;

		ws.mergeCells("C33:D33");
		ws.getCell("C33").alignment = textLeft;
		ws.getCell("C33").fill = toscaHeader;

		ws.getCell("F33").value = "";
		ws.getCell("F33").alignment = textLeft;
		ws.getCell("F33").fill = toscaHeader;

		ws.getCell("G33").value = "";
		ws.getCell("G33").alignment = textCenter;
		ws.getCell("G33").fill = toscaHeader;

		ws.mergeCells("H33:I33");
		ws.getCell("H33").alignment = textLeft;
		ws.getCell("H33").fill = toscaHeader;

		let datePrint = `${moment().format("YYYYMMDD")}_${moment().format(
			"HHmmss"
		)}`;
		const direct = "/tmp";
		await wb.xlsx
			.writeFile(`${direct}/VESSEL_DATA_${datePrint}.xlsx`)
			.then(() => {
				res.download(
					`${direct}/VESSEL_DATA_${datePrint}.xlsx`,
					`VESSEL_DATA_${datePrint}.xlsx`,
					(err) => {
						if (err) {
							console.log(err);
						} else {
							fs.unlinkSync(`${direct}/VESSEL_DATA_${datePrint}.xlsx`);
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
