const ExcelJS = require("exceljs");
const fs = require("fs");
const {
	borderBold,
	grayHeader,
	fontBold,
	textCenter,
	borderThin,
	textLeft,
	yellowHeader,
	textRight,
	fontBolder,
} = require("../utils/excel");
const moment = require("moment");
const path = require("path");

const exportExcel = async (req, res) => {
	try {
		const workbook = new ExcelJS.Workbook();
		const worksheet = workbook.addWorksheet("New Sheet", {
			properties: { tabColor: { argb: "FFC0000" } },
		});
		const direct = path.join(process.cwd(), "./public");
		// creating header column
		worksheet.mergeCells("A1:U1");
		worksheet.getCell("A1").value =
			"COAL TERMINAL SUMMARY REPORT 12 HOURS SHIFT";
		worksheet.getCell("A1").font = fontBolder;
		worksheet.getCell("A1").alignment = textCenter;

		worksheet.mergeCells("A3:I3");
		worksheet.mergeCells("J3:N3");
		worksheet.mergeCells("O3:Z3");
		worksheet.mergeCells("A4:A5");
		worksheet.mergeCells("A6:A7");
		worksheet.mergeCells("A8:A10");
		worksheet.mergeCells("B4:C5");
		worksheet.mergeCells("B6:C7");
		worksheet.mergeCells("B8:C10");
		worksheet.mergeCells("E4:F4");
		worksheet.mergeCells("E5:F5");
		worksheet.mergeCells("E6:F6");
		worksheet.mergeCells("E7:F7");
		worksheet.mergeCells("E8:F8");
		worksheet.mergeCells("E9:F9");
		worksheet.mergeCells("E10:F10");
		worksheet.mergeCells("H4:I4");
		worksheet.mergeCells("H5:I5");
		worksheet.mergeCells("H6:I6");
		worksheet.mergeCells("H7:I7");
		worksheet.mergeCells("H8:I8");
		worksheet.mergeCells("H9:I9");
		worksheet.mergeCells("H10:I10");
		worksheet.mergeCells("J5:J6");
		worksheet.mergeCells("J7:J8");
		worksheet.mergeCells("J9:M10");
		worksheet.mergeCells("K5:K6");
		worksheet.mergeCells("K7:K8");
		worksheet.mergeCells("L5:L6");
		worksheet.mergeCells("L7:L8");
		worksheet.mergeCells("M5:M6");
		worksheet.mergeCells("M7:M8");
		worksheet.mergeCells("N5:N6");
		worksheet.mergeCells("N7:N8");
		worksheet.mergeCells("N9:N10");
		worksheet.mergeCells("O4:P4");
		worksheet.mergeCells("Q4:R4");
		worksheet.mergeCells("O7:R8");
		worksheet.mergeCells("O9:R10");
		worksheet.mergeCells("S4:T4");
		worksheet.mergeCells("U4:V4");
		worksheet.mergeCells("S7:V8");
		worksheet.mergeCells("S9:V10");
		worksheet.mergeCells("W4:X4");
		worksheet.mergeCells("Y4:Z4");
		worksheet.mergeCells("A11:K11");
		worksheet.mergeCells("L11:V11");
		worksheet.mergeCells("A12:A13");
		worksheet.mergeCells("B12:C12");
		worksheet.mergeCells("D12:E12");
		worksheet.mergeCells("F12:G12");
		worksheet.mergeCells("H12:I12");
		worksheet.mergeCells("J12:K12");
		worksheet.mergeCells("L12:L13");
		worksheet.mergeCells("M12:N13");
		worksheet.mergeCells("O12:P12");
		worksheet.mergeCells("Q12:R12");
		worksheet.mergeCells("S12:T12");
		worksheet.mergeCells("U12:V12");

		worksheet.getCell("E3").value = "OPERATOR ON DUTY";
		worksheet.getCell("L3").value = "OLC WEIGHER DATA";
		worksheet.getCell("T3").value = "COAL SHIPPING";
		worksheet.getCell("A4").value = "DATE";
		worksheet.getCell("A6").value = "SHIFT";
		worksheet.getCell("A8").value = "CREW";
		worksheet.getCell("D4").value = "Supt";
		worksheet.getCell("D5").value = "Supv-in";
		worksheet.getCell("D6").value = "Supv-out";
		worksheet.getCell("D7").value = "C.Room 1";
		worksheet.getCell("D8").value = "C.Room 2";
		worksheet.getCell("D9").value = "Stacker 1";
		worksheet.getCell("D10").value = "Stacker 2";
		worksheet.getCell("G4").value = "Reclaimer 1";
		worksheet.getCell("G5").value = "Reclaimer 2";
		worksheet.getCell("G6").value = "NSL";
		worksheet.getCell("G7").value = "NTH DO";
		worksheet.getCell("G8").value = "SSL";
		worksheet.getCell("G9").value = "STH DO";
		worksheet.getCell("G10").value = "BLF";
		worksheet.getCell("J4").value = "Equipment";
		worksheet.getCell("K4").value = "Prima";
		worksheet.getCell("L4").value = "Pinang";
		worksheet.getCell("M4").value = "Melawan";
		worksheet.getCell("N4").value = "Total";
		worksheet.getCell("J5").value = "OLC 1";
		worksheet.getCell("J7").value = "OLC 2";
		worksheet.getCell("J9").value = "TOTAL COAL CONVEYED";
		worksheet.getCell("O4").value = "VESSEL 1 - MV :";
		worksheet.getCell("O5").value = "Reclaimer 1";
		worksheet.getCell("O7").value = "TOTAL BLF :";
		worksheet.getCell("O9").value =
			"TOTAL COAL SHIPPED (VESSEL 1 + VESSEL 2 + VESSEL 3 + BLF) :";
		worksheet.getCell("P5").value = "Reclaimer 2";
		worksheet.getCell("Q5").value = "Stamler";
		worksheet.getCell("R5").value = "Trestle";
		worksheet.getCell("S4").value = "VESSEL 2 - MV :";
		worksheet.getCell("S5").value = "Reclaimer 1";
		worksheet.getCell("T5").value = "Reclaimer 2";
		worksheet.getCell("U5").value = "Stamler";
		worksheet.getCell("V5").value = "Trestle";
		worksheet.getCell("W4").value = "VESSEL 3 - MV :";
		worksheet.getCell("W5").value = "Reclaimer 1";
		worksheet.getCell("X5").value = "Reclaimer 2";
		worksheet.getCell("Y5").value = "Stamler";
		worksheet.getCell("Z5").value = "Trestle";
		worksheet.getCell("A11").value = "COAL STACKING (INCOMING)";
		worksheet.getCell("L11").value = "COAL RECLAIMING (OUTGOING)";
		worksheet.getCell("A12").value = "Equipment";
		worksheet.getCell("B12").value = "Prima";
		worksheet.getCell("D12").value = "Pinang";
		worksheet.getCell("F12").value = "Melawan";
		worksheet.getCell("H12").value = "Position";
		worksheet.getCell("J12").value = "Total";
		worksheet.getCell("B13").value = "South";
		worksheet.getCell("C13").value = "North";
		worksheet.getCell("D13").value = "South";
		worksheet.getCell("E13").value = "North";
		worksheet.getCell("F13").value = "South";
		worksheet.getCell("G13").value = "North";
		worksheet.getCell("H13").value = "South";
		worksheet.getCell("I13").value = "North";
		worksheet.getCell("J13").value = "South";
		worksheet.getCell("K13").value = "North";
		worksheet.getCell("L12").value = "Equipment";
		worksheet.getCell("M12").value = "Vessel Name";
		worksheet.getCell("O12").value = "Prima";
		worksheet.getCell("Q12").value = "Pinang";
		worksheet.getCell("S12").value = "Melawan";
		worksheet.getCell("U12").value = "Total";
		worksheet.getCell("O13").value = "South";
		worksheet.getCell("P13").value = "North";
		worksheet.getCell("Q13").value = "South";
		worksheet.getCell("R13").value = "North";
		worksheet.getCell("S13").value = "South";
		worksheet.getCell("T13").value = "North";
		worksheet.getCell("U13").value = "South";
		worksheet.getCell("V13").value = "North";

		//cell with data from database
		worksheet.getCell("B4").value = "01/01/2024"; //date
		worksheet.getCell("B6").value = "shift malam"; // shift
		worksheet.getCell("B8").value = "Pak Yanuar"; // crew
		worksheet.getCell("E4").value = "Pak Yanuar"; // supv
		worksheet.getCell("E5").value = "Pak Yanuar"; // supv in
		worksheet.getCell("E6").value = "Pak Yanuar"; // supv out
		worksheet.getCell("E7").value = "Pak Yanuar"; // croom1
		worksheet.getCell("E8").value = "Pak Yanuar"; // croom2
		worksheet.getCell("E9").value = "Pak Yanuar"; // stacker1
		worksheet.getCell("E10").value = "Pak Yanuar"; // stacker2
		worksheet.getCell("H4").value = "Pak Yanuar"; // reclaimer1
		worksheet.getCell("H5").value = "Pak Yanuar"; // reclaimer2
		worksheet.getCell("H6").value = "Pak Yanuar"; // nsl
		worksheet.getCell("H7").value = "Pak Yanuar"; // nth do
		worksheet.getCell("H8").value = "Pak Yanuar"; // ssl
		worksheet.getCell("H9").value = "Pak Yanuar"; // sth do
		worksheet.getCell("H10").value = "Pak Yanuar"; // blf
		worksheet.getCell("K5").value = "Pak Yanuar"; // olc1 prima
		worksheet.getCell("K7").value = "Pak Yanuar"; // olc2 prima
		worksheet.getCell("L5").value = "Pak Yanuar"; // olc1 pinang
		worksheet.getCell("L7").value = "Pak Yanuar"; // olc2 pinang
		worksheet.getCell("M5").value = "Pak Yanuar"; // olc1 melawan
		worksheet.getCell("M7").value = "Pak Yanuar"; // olc2 melawan
		worksheet.getCell("N5").value = "Pak Yanuar"; // olc1 total
		worksheet.getCell("N7").value = "Pak Yanuar"; // olc2 total
		worksheet.getCell("N9").value = "Pak Yanuar"; // olc2 total
		worksheet.getCell("O6").value = "Pak Yanuar"; // coal reclaimer 1 vessel 1
		worksheet.getCell("P6").value = "Pak Yanuar"; // coal reclaimer 2 vessel 1
		worksheet.getCell("Q4").value = "Pak Yanuar"; // coal vessel 1 mv
		worksheet.getCell("Q6").value = "Pak Yanuar"; // coal stamler vessel 1
		worksheet.getCell("R6").value = "Pak Yanuar"; // coal trestle vessel 1

		worksheet.getCell("S6").value = "Pak Yanuar"; // coal reclaimer 1 vessel 2
		worksheet.getCell("T6").value = "Pak Yanuar"; // coal reclaimer 2 vessel 2
		worksheet.getCell("U4").value = "Pak Yanuar"; // coal vessel 2 mv
		worksheet.getCell("U6").value = "Pak Yanuar"; // coal stamler vessel 2
		worksheet.getCell("V6").value = "Pak Yanuar"; // coal trestle  vessel 2
		worksheet.getCell("S7").value = "Pak Yanuar"; // coal total blf
		worksheet.getCell("S9").value = "Pak Yanuar"; // coal coal shipped

		worksheet.getCell("W6").value = "Pak Yanuar"; // coal reclaimer 1 vessel 3
		worksheet.getCell("X6").value = "Pak Yanuar"; // coal reclaimer 2 vessel 3
		worksheet.getCell("Y4").value = "Pak Yanuar"; // coal vessel 3 mv
		worksheet.getCell("Y6").value = "Pak Yanuar"; // coal stamler vessel 3
		worksheet.getCell("Z6").value = "Pak Yanuar"; // coal trestle  vessel 3

		let data = [
			{
				a: "a14",
				b: "b14",
				c: "c14",
				d: "d14",
				e: "e14",
				f: "f14",
				g: "g14",
				h: "h14",
				i: "i14",
				j: "j14",
				k: "k14",
				l: "l14",
				m: "m14",
				n: "n14",
				o: "o14",
				p: "p14",
				q: "q14",
				r: "r14",
				s: "s14",
				t: "t14",
				u: "u14",
				v: "v14",
			},
			{
				a: "a15",
				b: "b15",
				c: "c15",
				d: "d15",
				e: "e15",
				f: "f15",
				g: "g15",
				h: "h15",
				i: "i15",
				j: "j15",
				k: "k15",
				l: "l15",
				m: "m15",
				n: "n15",
				o: "o15",
				p: "p15",
				q: "q15",
				r: "r15",
				s: "s15",
				t: "t15",
				u: "u15",
				v: "v15",
			},
			{
				a: "a16",
				b: "b16",
				c: "c16",
				d: "d16",
				e: "e16",
				f: "f16",
				g: "g16",
				h: "h16",
				i: "i16",
				j: "j16",
				k: "k16",
				l: "l16",
				m: "m16",
				n: "n16",
				o: "o16",
				p: "p16",
				q: "q16",
				r: "r16",
				s: "s16",
				t: "t16",
				u: "u16",
				v: "v16",
			},
			{
				a: "a17",
				b: "b17",
				c: "c17",
				d: "d17",
				e: "e17",
				f: "f17",
				g: "g17",
				h: "h17",
				i: "i17",
				j: "j17",
				k: "k17",
			},
			{
				a: "a18",
				b: "b18",
				c: "c18",
				d: "d18",
				e: "e18",
				f: "f18",
				g: "g18",
				h: "h18",
				i: "i18",
				j: "j18",
				k: "k18",
			},
		];
		let firstRow = 13;
		let rowCoalIncoming = 0;
		let rowCoalReclaiming = 0;
		for (let i in data) {
			firstRow++;
			worksheet.getCell("A" + firstRow).value = data[i].a;
			worksheet.getCell("A" + firstRow).alignment = textCenter;
			worksheet.getCell("A" + firstRow).border = borderThin;

			worksheet.getCell("B" + firstRow).value = data[i].b;
			worksheet.getCell("B" + firstRow).alignment = textCenter;
			worksheet.getCell("B" + firstRow).border = borderThin;

			worksheet.getCell("C" + firstRow).value = data[i].c;
			worksheet.getCell("C" + firstRow).alignment = textCenter;
			worksheet.getCell("C" + firstRow).border = borderThin;

			worksheet.getCell("D" + firstRow).value = data[i].d;
			worksheet.getCell("D" + firstRow).alignment = textCenter;
			worksheet.getCell("D" + firstRow).border = borderThin;

			worksheet.getCell("E" + firstRow).value = data[i].e;
			worksheet.getCell("E" + firstRow).alignment = textCenter;
			worksheet.getCell("E" + firstRow).border = borderThin;

			worksheet.getCell("F" + firstRow).value = data[i].f;
			worksheet.getCell("F" + firstRow).alignment = textCenter;
			worksheet.getCell("F" + firstRow).border = borderThin;

			worksheet.getCell("G" + firstRow).value = data[i].g;
			worksheet.getCell("G" + firstRow).alignment = textCenter;
			worksheet.getCell("G" + firstRow).border = borderThin;

			worksheet.getCell("H" + firstRow).value = data[i].h;
			worksheet.getCell("H" + firstRow).alignment = textCenter;
			worksheet.getCell("H" + firstRow).border = borderThin;

			worksheet.getCell("I" + firstRow).value = data[i].i;
			worksheet.getCell("I" + firstRow).alignment = textCenter;
			worksheet.getCell("I" + firstRow).border = borderThin;

			worksheet.getCell("J" + firstRow).value = data[i].j;
			worksheet.getCell("J" + firstRow).alignment = textCenter;
			worksheet.getCell("J" + firstRow).border = borderThin;

			worksheet.getCell("K" + firstRow).value = data[i].k;
			worksheet.getCell("K" + firstRow).alignment = textCenter;
			worksheet.getCell("K" + firstRow).border = borderThin;

			worksheet.getCell("L" + firstRow).value = data[i].l;
			worksheet.getCell("L" + firstRow).alignment = textCenter;
			worksheet.getCell("L" + firstRow).border = borderThin;

			if (data[i].l) {
				worksheet.mergeCells("M" + firstRow + ":" + "N" + firstRow);
			}

			worksheet.getCell("M" + firstRow).value = data[i].m;
			worksheet.getCell("M" + firstRow).alignment = textCenter;
			worksheet.getCell("M" + firstRow).border = borderThin;

			worksheet.getCell("O" + firstRow).value = data[i].o;
			worksheet.getCell("O" + firstRow).alignment = textCenter;
			worksheet.getCell("O" + firstRow).border = borderThin;

			worksheet.getCell("P" + firstRow).value = data[i].p;
			worksheet.getCell("P" + firstRow).alignment = textCenter;
			worksheet.getCell("P" + firstRow).border = borderThin;

			worksheet.getCell("Q" + firstRow).value = data[i].q;
			worksheet.getCell("Q" + firstRow).alignment = textCenter;
			worksheet.getCell("Q" + firstRow).border = borderThin;

			worksheet.getCell("R" + firstRow).value = data[i].r;
			worksheet.getCell("R" + firstRow).alignment = textCenter;
			worksheet.getCell("R" + firstRow).border = borderThin;

			worksheet.getCell("S" + firstRow).value = data[i].s;
			worksheet.getCell("S" + firstRow).alignment = textCenter;
			worksheet.getCell("S" + firstRow).border = borderThin;

			worksheet.getCell("T" + firstRow).value = data[i].t;
			worksheet.getCell("T" + firstRow).alignment = textCenter;
			worksheet.getCell("T" + firstRow).border = borderThin;

			worksheet.getCell("U" + firstRow).value = data[i].u;
			worksheet.getCell("U" + firstRow).alignment = textCenter;
			worksheet.getCell("U" + firstRow).border = borderThin;

			worksheet.getCell("V" + firstRow).value = data[i].v;
			worksheet.getCell("V" + firstRow).alignment = textCenter;
			worksheet.getCell("V" + firstRow).border = borderThin;

			if (data[i].l) {
				rowCoalReclaiming++;
			}
			//===========================================================================================
			worksheet.getCell("L" + (rowCoalReclaiming + 13 + 1)).value =
				"OLC#1 Bypass";
			worksheet.getCell("L" + (rowCoalReclaiming + 13 + 1)).border = borderThin;
			worksheet.getCell("L" + (rowCoalReclaiming + 13 + 1)).font = fontBold;
			worksheet.getCell("L" + (rowCoalReclaiming + 13 + 1)).alignment =
				textCenter;

			//===========================================================================================
			worksheet.getCell("L" + (rowCoalReclaiming + 13 + 2)).value =
				"OLC#1 Bypass";
			worksheet.getCell("L" + (rowCoalReclaiming + 13 + 2)).border = borderThin;
			worksheet.getCell("L" + (rowCoalReclaiming + 13 + 2)).font = fontBold;
			worksheet.getCell("L" + (rowCoalReclaiming + 13 + 2)).alignment =
				textCenter;
			//===========================================================================================
			worksheet.getCell("L" + (rowCoalReclaiming + 13 + 3)).value =
				"OLC#1 Bypass";
			worksheet.getCell("L" + (rowCoalReclaiming + 13 + 3)).border = borderThin;
			worksheet.getCell("L" + (rowCoalReclaiming + 13 + 3)).font = fontBold;
			worksheet.getCell("L" + (rowCoalReclaiming + 13 + 3)).alignment =
				textCenter;

			//===========================================================================================
			worksheet.getCell("L" + (rowCoalReclaiming + 13 + 4)).value =
				"OLC#2 Bypass";
			worksheet.getCell("L" + (rowCoalReclaiming + 13 + 4)).border = borderThin;
			worksheet.getCell("L" + (rowCoalReclaiming + 13 + 4)).font = fontBold;
			worksheet.getCell("L" + (rowCoalReclaiming + 13 + 4)).alignment =
				textCenter;

			//===========================================================================================
			worksheet.getCell("L" + (rowCoalReclaiming + 13 + 5)).value =
				"OLC#2 Bypass";
			worksheet.getCell("L" + (rowCoalReclaiming + 13 + 5)).border = borderThin;
			worksheet.getCell("L" + (rowCoalReclaiming + 13 + 5)).font = fontBold;
			worksheet.getCell("L" + (rowCoalReclaiming + 13 + 5)).alignment =
				textCenter;

			//===========================================================================================
			worksheet.getCell("L" + (rowCoalReclaiming + 13 + 6)).value =
				"OLC#2 Bypass";
			worksheet.getCell("L" + (rowCoalReclaiming + 13 + 6)).border = borderThin;
			worksheet.getCell("L" + (rowCoalReclaiming + 13 + 6)).font = fontBold;
			worksheet.getCell("L" + (rowCoalReclaiming + 13 + 6)).alignment =
				textCenter;

			//===========================================================================================

			rowCoalIncoming++;
		}
		let selisihRow = 6;
		for (let i = 0; selisihRow > i; selisihRow--) {
			if (selisihRow == 5) {
				worksheet.mergeCells(
					"A" +
						(rowCoalReclaiming + 13 + selisihRow) +
						":" +
						"I" +
						(rowCoalReclaiming + 13 + selisihRow)
				);
				worksheet.mergeCells(
					"J" +
						(rowCoalReclaiming + 13 + selisihRow) +
						":" +
						"K" +
						(rowCoalReclaiming + 13 + selisihRow)
				);
				worksheet.getCell("A" + (rowCoalReclaiming + 13 + selisihRow)).value =
					"TOTAL STACKING";
				worksheet.getCell("A" + (rowCoalReclaiming + 13 + selisihRow)).border =
					borderThin;
				worksheet.getCell(
					"A" + (rowCoalReclaiming + 13 + selisihRow)
				).alignment = textRight;
				worksheet.getCell("A" + (rowCoalReclaiming + 13 + selisihRow)).font =
					fontBold;
			} else if (selisihRow == 6) {
				worksheet.mergeCells(
					"A" +
						(rowCoalReclaiming + 13 + selisihRow) +
						":" +
						"K" +
						(rowCoalReclaiming + 13 + selisihRow)
				);
				worksheet.getCell("A" + (rowCoalReclaiming + 13 + selisihRow)).fill =
					yellowHeader;
				worksheet.getCell("A" + (rowCoalReclaiming + 13 + selisihRow)).border =
					borderThin;
			} else {
				worksheet.getCell(
					"A" + (rowCoalReclaiming + 13 + selisihRow)
				).alignment = textCenter;
				worksheet.getCell("A" + (rowCoalReclaiming + 13 + selisihRow)).border =
					borderThin;

				worksheet.getCell(
					"J" + (rowCoalReclaiming + 13 + selisihRow)
				).alignment = textCenter;
				worksheet.getCell("J" + (rowCoalReclaiming + 13 + selisihRow)).border =
					borderThin;
			}
			worksheet.getCell("B" + (rowCoalReclaiming + 13 + selisihRow)).alignment =
				textCenter;
			worksheet.getCell("B" + (rowCoalReclaiming + 13 + selisihRow)).border =
				borderThin;

			worksheet.getCell("C" + (rowCoalReclaiming + 13 + selisihRow)).alignment =
				textCenter;
			worksheet.getCell("C" + (rowCoalReclaiming + 13 + selisihRow)).border =
				borderThin;

			worksheet.getCell("D" + (rowCoalReclaiming + 13 + selisihRow)).alignment =
				textCenter;
			worksheet.getCell("D" + (rowCoalReclaiming + 13 + selisihRow)).border =
				borderThin;

			worksheet.getCell("E" + (rowCoalReclaiming + 13 + selisihRow)).alignment =
				textCenter;
			worksheet.getCell("E" + (rowCoalReclaiming + 13 + selisihRow)).border =
				borderThin;

			worksheet.getCell("F" + (rowCoalReclaiming + 13 + selisihRow)).alignment =
				textCenter;
			worksheet.getCell("F" + (rowCoalReclaiming + 13 + selisihRow)).border =
				borderThin;

			worksheet.getCell("G" + (rowCoalReclaiming + 13 + selisihRow)).alignment =
				textCenter;
			worksheet.getCell("G" + (rowCoalReclaiming + 13 + selisihRow)).border =
				borderThin;

			worksheet.getCell("H" + (rowCoalReclaiming + 13 + selisihRow)).alignment =
				textCenter;
			worksheet.getCell("H" + (rowCoalReclaiming + 13 + selisihRow)).border =
				borderThin;

			worksheet.getCell("I" + (rowCoalReclaiming + 13 + selisihRow)).alignment =
				textCenter;
			worksheet.getCell("I" + (rowCoalReclaiming + 13 + selisihRow)).border =
				borderThin;

			worksheet.getCell("K" + (rowCoalReclaiming + 13 + selisihRow)).alignment =
				textCenter;
			worksheet.getCell("K" + (rowCoalReclaiming + 13 + selisihRow)).border =
				borderThin;

			worksheet.getCell("M" + (rowCoalReclaiming + 13 + selisihRow)).alignment =
				textCenter;
			worksheet.getCell("M" + (rowCoalReclaiming + 13 + selisihRow)).border =
				borderThin;

			worksheet.getCell("N" + (rowCoalReclaiming + 13 + selisihRow)).alignment =
				textCenter;
			worksheet.getCell("N" + (rowCoalReclaiming + 13 + selisihRow)).border =
				borderThin;

			worksheet.getCell("O" + (rowCoalReclaiming + 13 + selisihRow)).alignment =
				textCenter;
			worksheet.getCell("O" + (rowCoalReclaiming + 13 + selisihRow)).border =
				borderThin;

			worksheet.getCell("P" + (rowCoalReclaiming + 13 + selisihRow)).alignment =
				textCenter;
			worksheet.getCell("P" + (rowCoalReclaiming + 13 + selisihRow)).border =
				borderThin;

			worksheet.getCell("Q" + (rowCoalReclaiming + 13 + selisihRow)).alignment =
				textCenter;
			worksheet.getCell("Q" + (rowCoalReclaiming + 13 + selisihRow)).border =
				borderThin;

			worksheet.getCell("R" + (rowCoalReclaiming + 13 + selisihRow)).alignment =
				textCenter;
			worksheet.getCell("R" + (rowCoalReclaiming + 13 + selisihRow)).border =
				borderThin;

			worksheet.getCell("S" + (rowCoalReclaiming + 13 + selisihRow)).alignment =
				textCenter;
			worksheet.getCell("S" + (rowCoalReclaiming + 13 + selisihRow)).border =
				borderThin;

			worksheet.getCell("T" + (rowCoalReclaiming + 13 + selisihRow)).alignment =
				textCenter;
			worksheet.getCell("T" + (rowCoalReclaiming + 13 + selisihRow)).border =
				borderThin;

			worksheet.getCell("U" + (rowCoalReclaiming + 13 + selisihRow)).alignment =
				textCenter;
			worksheet.getCell("U" + (rowCoalReclaiming + 13 + selisihRow)).border =
				borderThin;

			worksheet.getCell("V" + (rowCoalReclaiming + 13 + selisihRow)).alignment =
				textCenter;
			worksheet.getCell("V" + (rowCoalReclaiming + 13 + selisihRow)).border =
				borderThin;
		}
		worksheet.mergeCells(
			"L" +
				(rowCoalReclaiming + 13 + 7) +
				":" +
				"T" +
				(rowCoalReclaiming + 13 + 7)
		);
		worksheet.mergeCells(
			"A" +
				(rowCoalReclaiming + 13 + 7) +
				":" +
				"K" +
				(rowCoalReclaiming + 13 + 7)
		);
		worksheet.getCell("L" + (rowCoalReclaiming + 13 + 7)).value =
			"TOTAL RECLAIMING";
		worksheet.getCell("L" + (rowCoalReclaiming + 13 + 7)).border = borderThin;
		worksheet.getCell("L" + (rowCoalReclaiming + 13 + 7)).alignment = textRight;
		worksheet.getCell("L" + (rowCoalReclaiming + 13 + 7)).font = fontBold;

		worksheet.mergeCells(
			"U" +
				(rowCoalReclaiming + 13 + 7) +
				":" +
				"V" +
				(rowCoalReclaiming + 13 + 7)
		);
		worksheet.getCell("U" + (rowCoalReclaiming + 13 + 7)).border = borderThin;
		worksheet.getCell("U" + (rowCoalReclaiming + 13 + 7)).alignment = textRight;
		worksheet.getCell("U" + (rowCoalReclaiming + 13 + 7)).font = fontBold;

		worksheet.getCell("A" + (rowCoalReclaiming + 13 + 7)).border = borderThin;
		worksheet.getCell("A" + (rowCoalReclaiming + 13 + 7)).fill = yellowHeader;

		worksheet.mergeCells(
			"M" +
				(rowCoalReclaiming + 13 + 1) +
				":" +
				"N" +
				(rowCoalReclaiming + 13 + 1)
		);
		worksheet.mergeCells(
			"O" +
				(rowCoalReclaiming + 13 + 1) +
				":" +
				"P" +
				(rowCoalReclaiming + 13 + 1)
		);
		worksheet.mergeCells(
			"Q" +
				(rowCoalReclaiming + 13 + 1) +
				":" +
				"R" +
				(rowCoalReclaiming + 13 + 1)
		);
		worksheet.mergeCells(
			"S" +
				(rowCoalReclaiming + 13 + 1) +
				":" +
				"T" +
				(rowCoalReclaiming + 13 + 1)
		);
		worksheet.mergeCells(
			"U" +
				(rowCoalReclaiming + 13 + 1) +
				":" +
				"V" +
				(rowCoalReclaiming + 13 + 1)
		);
		worksheet.mergeCells(
			"M" +
				(rowCoalReclaiming + 13 + 2) +
				":" +
				"N" +
				(rowCoalReclaiming + 13 + 2)
		);
		worksheet.mergeCells(
			"O" +
				(rowCoalReclaiming + 13 + 2) +
				":" +
				"P" +
				(rowCoalReclaiming + 13 + 2)
		);
		worksheet.mergeCells(
			"Q" +
				(rowCoalReclaiming + 13 + 2) +
				":" +
				"R" +
				(rowCoalReclaiming + 13 + 2)
		);
		worksheet.mergeCells(
			"S" +
				(rowCoalReclaiming + 13 + 2) +
				":" +
				"T" +
				(rowCoalReclaiming + 13 + 2)
		);
		worksheet.mergeCells(
			"U" +
				(rowCoalReclaiming + 13 + 2) +
				":" +
				"V" +
				(rowCoalReclaiming + 13 + 2)
		);
		worksheet.mergeCells(
			"M" +
				(rowCoalReclaiming + 13 + 3) +
				":" +
				"N" +
				(rowCoalReclaiming + 13 + 3)
		);
		worksheet.mergeCells(
			"O" +
				(rowCoalReclaiming + 13 + 3) +
				":" +
				"P" +
				(rowCoalReclaiming + 13 + 3)
		);
		worksheet.mergeCells(
			"Q" +
				(rowCoalReclaiming + 13 + 3) +
				":" +
				"R" +
				(rowCoalReclaiming + 13 + 3)
		);
		worksheet.mergeCells(
			"S" +
				(rowCoalReclaiming + 13 + 3) +
				":" +
				"T" +
				(rowCoalReclaiming + 13 + 3)
		);
		worksheet.mergeCells(
			"U" +
				(rowCoalReclaiming + 13 + 3) +
				":" +
				"V" +
				(rowCoalReclaiming + 13 + 3)
		);
		worksheet.mergeCells(
			"M" +
				(rowCoalReclaiming + 13 + 4) +
				":" +
				"N" +
				(rowCoalReclaiming + 13 + 4)
		);
		worksheet.mergeCells(
			"O" +
				(rowCoalReclaiming + 13 + 4) +
				":" +
				"P" +
				(rowCoalReclaiming + 13 + 4)
		);
		worksheet.mergeCells(
			"Q" +
				(rowCoalReclaiming + 13 + 4) +
				":" +
				"R" +
				(rowCoalReclaiming + 13 + 4)
		);
		worksheet.mergeCells(
			"S" +
				(rowCoalReclaiming + 13 + 4) +
				":" +
				"T" +
				(rowCoalReclaiming + 13 + 4)
		);
		worksheet.mergeCells(
			"U" +
				(rowCoalReclaiming + 13 + 4) +
				":" +
				"V" +
				(rowCoalReclaiming + 13 + 4)
		);
		worksheet.mergeCells(
			"M" +
				(rowCoalReclaiming + 13 + 5) +
				":" +
				"N" +
				(rowCoalReclaiming + 13 + 5)
		);
		worksheet.mergeCells(
			"O" +
				(rowCoalReclaiming + 13 + 5) +
				":" +
				"P" +
				(rowCoalReclaiming + 13 + 5)
		);
		worksheet.mergeCells(
			"Q" +
				(rowCoalReclaiming + 13 + 5) +
				":" +
				"R" +
				(rowCoalReclaiming + 13 + 5)
		);
		worksheet.mergeCells(
			"S" +
				(rowCoalReclaiming + 13 + 5) +
				":" +
				"T" +
				(rowCoalReclaiming + 13 + 5)
		);
		worksheet.mergeCells(
			"U" +
				(rowCoalReclaiming + 13 + 5) +
				":" +
				"V" +
				(rowCoalReclaiming + 13 + 5)
		);
		worksheet.mergeCells(
			"M" +
				(rowCoalReclaiming + 13 + 6) +
				":" +
				"N" +
				(rowCoalReclaiming + 13 + 6)
		);
		worksheet.mergeCells(
			"O" +
				(rowCoalReclaiming + 13 + 6) +
				":" +
				"P" +
				(rowCoalReclaiming + 13 + 6)
		);
		worksheet.mergeCells(
			"Q" +
				(rowCoalReclaiming + 13 + 6) +
				":" +
				"R" +
				(rowCoalReclaiming + 13 + 6)
		);
		worksheet.mergeCells(
			"S" +
				(rowCoalReclaiming + 13 + 6) +
				":" +
				"T" +
				(rowCoalReclaiming + 13 + 6)
		);
		worksheet.mergeCells(
			"U" +
				(rowCoalReclaiming + 13 + 6) +
				":" +
				"V" +
				(rowCoalReclaiming + 13 + 6)
		);

		//Header BARGE LOADING SHIPPING DETAIL
		worksheet.mergeCells(
			"A" +
				(rowCoalReclaiming + 13 + 8) +
				":" +
				"V" +
				(rowCoalReclaiming + 13 + 8)
		);
		worksheet.getCell("A" + (rowCoalReclaiming + 13 + 8)).value =
			"BARGE LOADING SHIPPING DETAIL";
		worksheet.getCell("A" + (rowCoalReclaiming + 13 + 8)).alignment =
			textCenter;
		worksheet.getCell("A" + (rowCoalReclaiming + 13 + 8)).fill = grayHeader;
		worksheet.getCell("A" + (rowCoalReclaiming + 13 + 8)).border = borderBold;
		worksheet.getCell("A" + (rowCoalReclaiming + 13 + 8)).font = fontBold;

		//Barge
		worksheet.mergeCells(
			"A" +
				(rowCoalReclaiming + 13 + 9) +
				":" +
				"A" +
				(rowCoalReclaiming + 13 + 10)
		);
		worksheet.getCell("A" + (rowCoalReclaiming + 13 + 9)).value = "No.";
		worksheet.getCell("A" + (rowCoalReclaiming + 13 + 9)).font = fontBold;
		worksheet.getCell("A" + (rowCoalReclaiming + 13 + 9)).alignment =
			textCenter;
		worksheet.getCell("A" + (rowCoalReclaiming + 13 + 9)).border = borderThin;

		worksheet.mergeCells(
			"B" +
				(rowCoalReclaiming + 13 + 9) +
				":" +
				"D" +
				(rowCoalReclaiming + 13 + 10)
		);
		worksheet.getCell("B" + (rowCoalReclaiming + 13 + 9)).value = "Vessel";
		worksheet.getCell("B" + (rowCoalReclaiming + 13 + 9)).font = fontBold;
		worksheet.getCell("B" + (rowCoalReclaiming + 13 + 9)).alignment =
			textCenter;
		worksheet.getCell("B" + (rowCoalReclaiming + 13 + 9)).border = borderThin;

		worksheet.mergeCells(
			"E" +
				(rowCoalReclaiming + 13 + 9) +
				":" +
				"N" +
				(rowCoalReclaiming + 13 + 9)
		);
		worksheet.getCell("E" + (rowCoalReclaiming + 13 + 9)).value = "Barge";
		worksheet.getCell("E" + (rowCoalReclaiming + 13 + 9)).font = fontBold;
		worksheet.getCell("E" + (rowCoalReclaiming + 13 + 9)).alignment =
			textCenter;
		worksheet.getCell("E" + (rowCoalReclaiming + 13 + 9)).border = borderThin;

		worksheet.mergeCells(
			"E" +
				(rowCoalReclaiming + 13 + 10) +
				":" +
				"F" +
				(rowCoalReclaiming + 13 + 10)
		);
		worksheet.getCell("E" + (rowCoalReclaiming + 13 + 10)).value = "Name";
		worksheet.getCell("E" + (rowCoalReclaiming + 13 + 10)).style = {
			font: fontBold,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("G" + (rowCoalReclaiming + 13 + 10)).value =
			"Along Side (First Line)";
		worksheet.getCell("G" + (rowCoalReclaiming + 13 + 9)).font = fontBold;
		worksheet.getCell("G" + (rowCoalReclaiming + 13 + 9)).alignment =
			textCenter;
		worksheet.getCell("G" + (rowCoalReclaiming + 13 + 9)).border = borderThin;

		worksheet.mergeCells(
			"H" +
				(rowCoalReclaiming + 13 + 10) +
				":" +
				"I" +
				(rowCoalReclaiming + 13 + 10)
		);
		worksheet.getCell("H" + (rowCoalReclaiming + 13 + 10)).value =
			"Commenced Loading";
		worksheet.getCell("H" + (rowCoalReclaiming + 13 + 9)).font = fontBold;
		worksheet.getCell("H" + (rowCoalReclaiming + 13 + 9)).alignment =
			textCenter;
		worksheet.getCell("H" + (rowCoalReclaiming + 13 + 9)).border = borderThin;

		worksheet.mergeCells(
			"J" +
				(rowCoalReclaiming + 13 + 10) +
				":" +
				"K" +
				(rowCoalReclaiming + 13 + 10)
		);
		worksheet.getCell("J" + (rowCoalReclaiming + 13 + 10)).value =
			"Completed Loading";
		worksheet.getCell("J" + (rowCoalReclaiming + 13 + 9)).font = fontBold;
		worksheet.getCell("J" + (rowCoalReclaiming + 13 + 9)).alignment =
			textCenter;
		worksheet.getCell("J" + (rowCoalReclaiming + 13 + 9)).border = borderThin;

		worksheet.getCell("L" + (rowCoalReclaiming + 13 + 10)).value =
			"Cast Off (Last Line)";
		worksheet.getCell("L" + (rowCoalReclaiming + 13 + 9)).font = fontBold;
		worksheet.getCell("L" + (rowCoalReclaiming + 13 + 9)).alignment =
			textCenter;
		worksheet.getCell("L" + (rowCoalReclaiming + 13 + 9)).border = borderThin;

		worksheet.mergeCells(
			"M" +
				(rowCoalReclaiming + 13 + 10) +
				":" +
				"N" +
				(rowCoalReclaiming + 13 + 10)
		);
		worksheet.getCell("M" + (rowCoalReclaiming + 13 + 10)).value =
			"Total Loading Time";
		worksheet.getCell("M" + (rowCoalReclaiming + 13 + 9)).font = fontBold;
		worksheet.getCell("M" + (rowCoalReclaiming + 13 + 9)).alignment =
			textCenter;
		worksheet.getCell("M" + (rowCoalReclaiming + 13 + 9)).border = borderThin;

		let dataBarge = [
			{
				a: "a16",
				b: "b16",
				e: "e16",
				g: "g16",
				h: "h16",
				j: "j16",
				l: "l16",
				m: "m16",
				o: "o16",
				p: "p16",
				q: "q16",
				r: "r16",
				s: "s16",
				u: "u16",
			},
			{
				a: "a16",
				b: "b16",
				e: "e16",
				g: "g16",
				h: "h16",
				j: "j16",
				l: "l16",
				m: "m16",
				o: "o16",
				p: "p16",
				q: "q16",
				r: "r16",
				s: "s16",
				u: "u16",
			},
			{
				a: "a16",
				b: "b16",
				e: "e16",
				g: "g16",
				h: "h16",
				j: "j16",
				l: "l16",
				m: "m16",
				o: "o16",
				p: "p16",
				q: "q16",
				r: "r16",
				s: "s16",
				u: "u16",
			},
		];
		var startBargeLoading = rowCoalReclaiming + 13 + 10;

		for (let i in dataBarge) {
			startBargeLoading++;
			worksheet.getCell("A" + startBargeLoading).value = dataBarge[i].a;
			worksheet.getCell("A" + startBargeLoading).fill = yellowHeader;
			worksheet.getCell("A" + startBargeLoading).alignment = textCenter;
			worksheet.getCell("A" + startBargeLoading).border = borderThin;

			worksheet.mergeCells(
				"B" + startBargeLoading + ":" + "D" + startBargeLoading
			);
			worksheet.getCell("B" + startBargeLoading).value = dataBarge[i].b;
			worksheet.getCell("B" + startBargeLoading).fill = yellowHeader;
			worksheet.getCell("B" + startBargeLoading).alignment = textCenter;
			worksheet.getCell("B" + startBargeLoading).border = borderThin;

			worksheet.mergeCells(
				"E" + startBargeLoading + ":" + "F" + startBargeLoading
			);
			worksheet.getCell("E" + startBargeLoading).value = dataBarge[i].e;
			worksheet.getCell("E" + startBargeLoading).fill = yellowHeader;
			worksheet.getCell("E" + startBargeLoading).alignment = textCenter;
			worksheet.getCell("E" + startBargeLoading).border = borderThin;

			worksheet.getCell("G" + startBargeLoading).value = dataBarge[i].g;
			worksheet.getCell("G" + startBargeLoading).fill = yellowHeader;
			worksheet.getCell("G" + startBargeLoading).alignment = textCenter;
			worksheet.getCell("G" + startBargeLoading).border = borderThin;

			worksheet.mergeCells(
				"H" + startBargeLoading + ":" + "I" + startBargeLoading
			);
			worksheet.getCell("H" + startBargeLoading).value = dataBarge[i].h;
			worksheet.getCell("H" + startBargeLoading).fill = yellowHeader;
			worksheet.getCell("H" + startBargeLoading).alignment = textCenter;
			worksheet.getCell("H" + startBargeLoading).border = borderThin;

			worksheet.mergeCells(
				"J" + startBargeLoading + ":" + "K" + startBargeLoading
			);
			worksheet.getCell("J" + startBargeLoading).value = dataBarge[i].j;
			worksheet.getCell("J" + startBargeLoading).fill = yellowHeader;
			worksheet.getCell("J" + startBargeLoading).alignment = textCenter;
			worksheet.getCell("J" + startBargeLoading).border = borderThin;

			worksheet.getCell("L" + startBargeLoading).value = dataBarge[i].l;
			worksheet.getCell("L" + startBargeLoading).fill = yellowHeader;
			worksheet.getCell("L" + startBargeLoading).alignment = textCenter;
			worksheet.getCell("L" + startBargeLoading).border = borderThin;

			worksheet.mergeCells(
				"M" + startBargeLoading + ":" + "N" + startBargeLoading
			);
			worksheet.getCell("M" + startBargeLoading).value = dataBarge[i].m;
			worksheet.getCell("M" + startBargeLoading).fill = yellowHeader;
			worksheet.getCell("M" + startBargeLoading).alignment = textCenter;
			worksheet.getCell("M" + startBargeLoading).border = borderThin;

			worksheet.getCell("O" + startBargeLoading).value = dataBarge[i].o;
			worksheet.getCell("O" + startBargeLoading).fill = yellowHeader;
			worksheet.getCell("O" + startBargeLoading).alignment = textCenter;
			worksheet.getCell("O" + startBargeLoading).border = borderThin;

			worksheet.getCell("P" + startBargeLoading).value = dataBarge[i].p;
			worksheet.getCell("P" + startBargeLoading).fill = yellowHeader;
			worksheet.getCell("P" + startBargeLoading).alignment = textCenter;
			worksheet.getCell("P" + startBargeLoading).border = borderThin;

			worksheet.getCell("Q" + startBargeLoading).value = dataBarge[i].q;
			worksheet.getCell("Q" + startBargeLoading).fill = yellowHeader;
			worksheet.getCell("Q" + startBargeLoading).alignment = textCenter;
			worksheet.getCell("Q" + startBargeLoading).border = borderThin;

			worksheet.getCell("R" + startBargeLoading).value = dataBarge[i].r;
			worksheet.getCell("R" + startBargeLoading).fill = yellowHeader;
			worksheet.getCell("R" + startBargeLoading).alignment = textCenter;
			worksheet.getCell("R" + startBargeLoading).border = borderThin;

			worksheet.mergeCells(
				"S" + startBargeLoading + ":" + "T" + startBargeLoading
			);
			worksheet.getCell("S" + startBargeLoading).value = dataBarge[i].s;
			worksheet.getCell("S" + startBargeLoading).fill = yellowHeader;
			worksheet.getCell("S" + startBargeLoading).alignment = textCenter;
			worksheet.getCell("S" + startBargeLoading).border = borderThin;

			worksheet.mergeCells(
				"U" + startBargeLoading + ":" + "V" + startBargeLoading
			);
			worksheet.getCell("U" + startBargeLoading).value = dataBarge[i].u;
			worksheet.getCell("U" + startBargeLoading).fill = yellowHeader;
			worksheet.getCell("U" + startBargeLoading).alignment = textCenter;
			worksheet.getCell("U" + startBargeLoading).border = borderThin;
		}

		//Weigher Data
		worksheet.mergeCells(
			"O" +
				(rowCoalReclaiming + 13 + 9) +
				":" +
				"R" +
				(rowCoalReclaiming + 13 + 9)
		);
		worksheet.getCell("O" + (rowCoalReclaiming + 13 + 9)).value =
			"Weigher Data";
		worksheet.getCell("O" + (rowCoalReclaiming + 13 + 9)).font = fontBold;
		worksheet.getCell("O" + (rowCoalReclaiming + 13 + 9)).alignment =
			textCenter;
		worksheet.getCell("O" + (rowCoalReclaiming + 13 + 9)).border = borderThin;

		worksheet.getCell("O" + (rowCoalReclaiming + 13 + 10)).value = "Quality";
		worksheet.getCell("O" + (rowCoalReclaiming + 13 + 10)).font = fontBold;
		worksheet.getCell("O" + (rowCoalReclaiming + 13 + 10)).alignment =
			textCenter;
		worksheet.getCell("O" + (rowCoalReclaiming + 13 + 10)).border = borderThin;

		worksheet.getCell("P" + (rowCoalReclaiming + 13 + 10)).value = "Start";
		worksheet.getCell("P" + (rowCoalReclaiming + 13 + 10)).font = fontBold;
		worksheet.getCell("P" + (rowCoalReclaiming + 13 + 10)).alignment =
			textCenter;
		worksheet.getCell("P" + (rowCoalReclaiming + 13 + 10)).border = borderThin;

		worksheet.getCell("Q" + (rowCoalReclaiming + 13 + 10)).value = "Stop";
		worksheet.getCell("Q" + (rowCoalReclaiming + 13 + 10)).font = fontBold;
		worksheet.getCell("Q" + (rowCoalReclaiming + 13 + 10)).alignment =
			textCenter;
		worksheet.getCell("Q" + (rowCoalReclaiming + 13 + 10)).border = borderThin;

		worksheet.getCell("R" + (rowCoalReclaiming + 13 + 10)).value = "Total";
		worksheet.getCell("R" + (rowCoalReclaiming + 13 + 10)).font = fontBold;
		worksheet.getCell("R" + (rowCoalReclaiming + 13 + 10)).alignment =
			textCenter;
		worksheet.getCell("R" + (rowCoalReclaiming + 13 + 10)).border = borderThin;

		worksheet.mergeCells(
			"S" +
				(rowCoalReclaiming + 13 + 9) +
				":" +
				"T" +
				(rowCoalReclaiming + 13 + 10)
		);
		worksheet.getCell("S" + (rowCoalReclaiming + 13 + 9)).value =
			"Tonnes Draft";
		worksheet.getCell("S" + (rowCoalReclaiming + 13 + 9)).font = fontBold;
		worksheet.getCell("S" + (rowCoalReclaiming + 13 + 9)).alignment =
			textCenter;
		worksheet.getCell("S" + (rowCoalReclaiming + 13 + 9)).border = borderThin;

		worksheet.mergeCells(
			"U" +
				(rowCoalReclaiming + 13 + 9) +
				":" +
				"V" +
				(rowCoalReclaiming + 13 + 10)
		);
		worksheet.getCell("U" + (rowCoalReclaiming + 13 + 9)).value =
			"Tonnes MCC Recorded";
		worksheet.getCell("U" + (rowCoalReclaiming + 13 + 9)).font = fontBold;
		worksheet.getCell("U" + (rowCoalReclaiming + 13 + 9)).alignment =
			textCenter;
		worksheet.getCell("U" + (rowCoalReclaiming + 13 + 9)).border = borderThin;

		//Equpment Detail
		let startEquipmentDetail = startBargeLoading + 1;
		let secondEquipmentDetail = startEquipmentDetail + 1;

		worksheet.mergeCells(
			"A" + startEquipmentDetail + ":" + "V" + startEquipmentDetail
		);
		worksheet.getCell("A" + startEquipmentDetail).value = "EQUIPMENT DETAIL";
		worksheet.getCell("A" + startEquipmentDetail).fill = grayHeader;
		worksheet.getCell("A" + startEquipmentDetail).alignment = textCenter;
		worksheet.getCell("A" + startEquipmentDetail).border = borderThin;
		worksheet.getCell("A" + startEquipmentDetail).font = fontBold;

		worksheet.getCell("A" + secondEquipmentDetail).value = "Equipment";
		worksheet.getCell("A" + secondEquipmentDetail).alignment = textCenter;
		worksheet.getCell("A" + secondEquipmentDetail).border = borderThin;
		worksheet.getCell("A" + secondEquipmentDetail).font = fontBold;

		worksheet.mergeCells(
			"B" + secondEquipmentDetail + ":" + "C" + secondEquipmentDetail
		);
		worksheet.getCell("B" + secondEquipmentDetail).value = "Operator";
		worksheet.getCell("B" + secondEquipmentDetail).alignment = textCenter;
		worksheet.getCell("B" + secondEquipmentDetail).border = borderThin;
		worksheet.getCell("B" + secondEquipmentDetail).font = fontBold;

		worksheet.mergeCells(
			"D" + secondEquipmentDetail + ":" + "E" + secondEquipmentDetail
		);
		worksheet.getCell("D" + secondEquipmentDetail).value = "Fuel";
		worksheet.getCell("D" + secondEquipmentDetail).alignment = textCenter;
		worksheet.getCell("D" + secondEquipmentDetail).border = borderThin;
		worksheet.getCell("D" + secondEquipmentDetail).font = fontBold;

		worksheet.mergeCells(
			"F" + secondEquipmentDetail + ":" + "G" + secondEquipmentDetail
		);
		worksheet.getCell("F" + secondEquipmentDetail).value = "Start";
		worksheet.getCell("F" + secondEquipmentDetail).alignment = textCenter;
		worksheet.getCell("F" + secondEquipmentDetail).border = borderThin;
		worksheet.getCell("F" + secondEquipmentDetail).font = fontBold;

		worksheet.mergeCells(
			"H" + secondEquipmentDetail + ":" + "I" + secondEquipmentDetail
		);
		worksheet.getCell("H" + secondEquipmentDetail).value = "Stop";
		worksheet.getCell("H" + secondEquipmentDetail).alignment = textCenter;
		worksheet.getCell("H" + secondEquipmentDetail).border = borderThin;
		worksheet.getCell("H" + secondEquipmentDetail).font = fontBold;

		worksheet.mergeCells(
			"J" + secondEquipmentDetail + ":" + "K" + secondEquipmentDetail
		);
		worksheet.getCell("J" + secondEquipmentDetail).value = "Total Hours";
		worksheet.getCell("J" + secondEquipmentDetail).alignment = textCenter;
		worksheet.getCell("J" + secondEquipmentDetail).border = borderThin;
		worksheet.getCell("J" + secondEquipmentDetail).font = fontBold;

		worksheet.getCell("L" + secondEquipmentDetail).value = "Equipment";
		worksheet.getCell("L" + secondEquipmentDetail).alignment = textCenter;
		worksheet.getCell("L" + secondEquipmentDetail).border = borderThin;
		worksheet.getCell("L" + secondEquipmentDetail).font = fontBold;

		worksheet.mergeCells(
			"M" + secondEquipmentDetail + ":" + "N" + secondEquipmentDetail
		);
		worksheet.getCell("M" + secondEquipmentDetail).value = "Operator";
		worksheet.getCell("M" + secondEquipmentDetail).alignment = textCenter;
		worksheet.getCell("M" + secondEquipmentDetail).border = borderThin;
		worksheet.getCell("M" + secondEquipmentDetail).font = fontBold;

		worksheet.mergeCells(
			"O" + secondEquipmentDetail + ":" + "P" + secondEquipmentDetail
		);
		worksheet.getCell("O" + secondEquipmentDetail).value = "Fuel";
		worksheet.getCell("O" + secondEquipmentDetail).alignment = textCenter;
		worksheet.getCell("O" + secondEquipmentDetail).border = borderThin;
		worksheet.getCell("O" + secondEquipmentDetail).font = fontBold;

		worksheet.mergeCells(
			"Q" + secondEquipmentDetail + ":" + "R" + secondEquipmentDetail
		);
		worksheet.getCell("Q" + secondEquipmentDetail).value = "Start";
		worksheet.getCell("Q" + secondEquipmentDetail).alignment = textCenter;
		worksheet.getCell("Q" + secondEquipmentDetail).border = borderThin;
		worksheet.getCell("Q" + secondEquipmentDetail).font = fontBold;

		worksheet.mergeCells(
			"S" + secondEquipmentDetail + ":" + "T" + secondEquipmentDetail
		);
		worksheet.getCell("S" + secondEquipmentDetail).value = "Stop";
		worksheet.getCell("S" + secondEquipmentDetail).alignment = textCenter;
		worksheet.getCell("S" + secondEquipmentDetail).border = borderThin;
		worksheet.getCell("S" + secondEquipmentDetail).font = fontBold;

		worksheet.mergeCells(
			"U" + secondEquipmentDetail + ":" + "V" + secondEquipmentDetail
		);
		worksheet.getCell("U" + secondEquipmentDetail).value = "Total Hours";
		worksheet.getCell("U" + secondEquipmentDetail).alignment = textCenter;
		worksheet.getCell("U" + secondEquipmentDetail).border = borderThin;
		worksheet.getCell("U" + secondEquipmentDetail).font = fontBold;

		//eqipment Detail data
		let dataEquipmentDetail = [
			{
				a: "DOZER E529",
				b: "bed",
				d: "ded",
				f: "fed",
				h: "hed",
				j: "jed",
				l: "led",
				m: "med",
				o: "oed",
				q: "qed",
				s: "sed",
				u: "ued",
			},
			{
				a: "DOZER E530",
				b: "bed",
				d: "ded",
				f: "fed",
				h: "hed",
				j: "jed",
				l: "led",
				m: "med",
				o: "oed",
				q: "qed",
				s: "sed",
				u: "ued",
			},
			{
				a: "DOZER E532",
				b: "bed",
				d: "ded",
				f: "fed",
				h: "hed",
				j: "jed",
				l: "led",
				m: "med",
				o: "oed",
				q: "qed",
				s: "sed",
				u: "ued",
			},
			{
				a: "DOZER E557",
				b: "bed",
				d: "ded",
				f: "fed",
				h: "hed",
				j: "jed",
				l: "led",
				m: "med",
				o: "oed",
				q: "qed",
				s: "sed",
				u: "ued",
			},
			{
				a: "DOZER E557",
				b: "bed",
				d: "ded",
				f: "fed",
				h: "hed",
				j: "jed",
				l: "led",
				m: "med",
				o: "oed",
				q: "qed",
				s: "sed",
				u: "ued",
			},
			{
				a: "DOZER E576",
				b: "bed",
				d: "ded",
				f: "fed",
				h: "hed",
				j: "jed",
				l: "led",
				m: "med",
				o: "oed",
				q: "qed",
				s: "sed",
				u: "ued",
			},
		];

		let startEquipmentDetailData = secondEquipmentDetail;
		for (let i in dataEquipmentDetail) {
			startEquipmentDetailData++;
			worksheet.getCell("A" + startEquipmentDetailData).value =
				dataEquipmentDetail[i].a;
			worksheet.getCell("A" + startEquipmentDetailData).border = borderThin;
			worksheet.getCell("A" + startEquipmentDetailData).fill = yellowHeader;

			worksheet.mergeCells(
				"B" + startEquipmentDetailData + ":" + "C" + startEquipmentDetailData
			);
			worksheet.getCell("B" + startEquipmentDetailData).value =
				dataEquipmentDetail[i].b;
			worksheet.getCell("B" + startEquipmentDetailData).border = borderThin;
			worksheet.getCell("B" + startEquipmentDetailData).fill = yellowHeader;
			worksheet.getCell("B" + startEquipmentDetailData).alignment = textCenter;

			worksheet.mergeCells(
				"D" + startEquipmentDetailData + ":" + "E" + startEquipmentDetailData
			);
			worksheet.getCell("D" + startEquipmentDetailData).value =
				dataEquipmentDetail[i].d;
			worksheet.getCell(
				"D" + startEquipmentDetailData + ":" + "E" + startEquipmentDetailData
			).style = {
				border: borderThin,
				fill: yellowHeader,
				alignment: textCenter,
			};
			worksheet.mergeCells(
				"F" + startEquipmentDetailData + ":" + "G" + startEquipmentDetailData
			);
			worksheet.getCell("F" + startEquipmentDetailData).value =
				dataEquipmentDetail[i].f;
			worksheet.getCell("F" + startEquipmentDetailData).style = {
				border: borderThin,
				fill: yellowHeader,
				alignment: textCenter,
			};

			worksheet.mergeCells(
				"H" + startEquipmentDetailData + ":" + "I" + startEquipmentDetailData
			);
			worksheet.getCell("H" + startEquipmentDetailData).value =
				dataEquipmentDetail[i].h;
			worksheet.getCell("H" + startEquipmentDetailData).style = {
				border: borderThin,
				fill: yellowHeader,
				alignment: textCenter,
			};
			worksheet.mergeCells(
				"J" + startEquipmentDetailData + ":" + "K" + startEquipmentDetailData
			);
			worksheet.getCell("J" + startEquipmentDetailData).value =
				dataEquipmentDetail[i].j;
			worksheet.getCell("J" + startEquipmentDetailData).style = {
				border: borderThin,
				fill: yellowHeader,
				alignment: textCenter,
			};
			worksheet.getCell("L" + startEquipmentDetailData).value =
				dataEquipmentDetail[i].l;
			worksheet.getCell("L" + startEquipmentDetailData).style = {
				border: borderThin,
				fill: yellowHeader,
				alignment: textCenter,
			};
			worksheet.mergeCells(
				"M" + startEquipmentDetailData + ":" + "N" + startEquipmentDetailData
			);
			worksheet.getCell("M" + startEquipmentDetailData).value =
				dataEquipmentDetail[i].m;
			worksheet.getCell("M" + startEquipmentDetailData).style = {
				border: borderThin,
				fill: yellowHeader,
				alignment: textCenter,
			};
			worksheet.mergeCells(
				"O" + startEquipmentDetailData + ":" + "P" + startEquipmentDetailData
			);
			worksheet.getCell("O" + startEquipmentDetailData).value =
				dataEquipmentDetail[i].o;
			worksheet.getCell("O" + startEquipmentDetailData).style = {
				border: borderThin,
				fill: yellowHeader,
				alignment: textCenter,
			};
			worksheet.mergeCells(
				"Q" + startEquipmentDetailData + ":" + "R" + startEquipmentDetailData
			);
			worksheet.getCell("Q" + startEquipmentDetailData).value =
				dataEquipmentDetail[i].q;
			worksheet.getCell("Q" + startEquipmentDetailData).style = {
				border: borderThin,
				fill: yellowHeader,
				alignment: textCenter,
			};
			worksheet.mergeCells(
				"S" + startEquipmentDetailData + ":" + "T" + startEquipmentDetailData
			);
			worksheet.getCell("S" + startEquipmentDetailData).value =
				dataEquipmentDetail[i].s;
			worksheet.getCell("S" + startEquipmentDetailData).style = {
				border: borderThin,
				fill: yellowHeader,
				alignment: textCenter,
			};
			worksheet.mergeCells(
				"U" + startEquipmentDetailData + ":" + "V" + startEquipmentDetailData
			);
			worksheet.getCell("U" + startEquipmentDetailData).value =
				dataEquipmentDetail[i].u;
			worksheet.getCell("U" + startEquipmentDetailData).style = {
				border: borderThin,
				fill: yellowHeader,
				alignment: textCenter,
			};
		}
		// Notes
		let startNotes = startEquipmentDetailData + 1;
		worksheet.mergeCells("A" + startNotes + ":" + "V" + startNotes);
		worksheet.getCell("A" + startNotes).value = "Notes";
		worksheet.getCell("A" + startNotes).style = {
			border: borderBold,
			fill: grayHeader,
			alignment: textCenter,
		};
		worksheet.mergeCells(
			"A" + (startNotes + 1) + ":" + "V" + (startNotes + 10)
		);
		worksheet.getCell("A" + (startNotes + 1)).value =
			"ini notes nya summary report";
		worksheet.getCell("A" + (startNotes + 1)).style = {
			border: borderThin,
			fill: yellowHeader,
			alignment: textLeft,
		};

		//style column

		worksheet.getCell("E3").fill = grayHeader;
		worksheet.getCell("E3").font = fontBold;
		worksheet.getCell("E3").alignment = textCenter;
		worksheet.getCell("E3").border = borderBold;

		worksheet.getCell("L3").fill = grayHeader;
		worksheet.getCell("L3").font = fontBold;
		worksheet.getCell("L3").alignment = textCenter;
		worksheet.getCell("L3").border = borderBold;

		worksheet.getCell("T3").fill = grayHeader;
		worksheet.getCell("T3").font = fontBold;
		worksheet.getCell("T3").alignment = textCenter;
		worksheet.getCell("T3").border = borderBold;

		worksheet.getCell("A4").fill = grayHeader;
		worksheet.getCell("A4").font = fontBold;
		worksheet.getCell("A4").alignment = textCenter;
		worksheet.getCell("A4").border = borderThin;

		worksheet.getCell("A6").fill = grayHeader;
		worksheet.getCell("A6").font = fontBold;
		worksheet.getCell("A6").alignment = textCenter;
		worksheet.getCell("A6").border = borderThin;

		worksheet.getCell("A8").fill = grayHeader;
		worksheet.getCell("A8").font = fontBold;
		worksheet.getCell("A8").alignment = textCenter;
		worksheet.getCell("A8").border = borderThin;

		// column for title
		worksheet.getCell("D4").style = {
			font: fontBold,
			alignment: textLeft,
			border: borderThin,
		};
		worksheet.getCell("D5").style = {
			font: fontBold,
			alignment: textLeft,
			border: borderThin,
		};
		worksheet.getCell("D6").style = {
			font: fontBold,
			alignment: textLeft,
			border: borderThin,
		};
		worksheet.getCell("D7").style = {
			font: fontBold,
			alignment: textLeft,
			border: borderThin,
		};
		worksheet.getCell("D8").style = {
			font: fontBold,
			alignment: textLeft,
			border: borderThin,
		};
		worksheet.getCell("D9").style = {
			font: fontBold,
			alignment: textLeft,
			border: borderThin,
		};
		worksheet.getCell("D10").style = {
			font: fontBold,
			alignment: textLeft,
			border: borderThin,
		};
		worksheet.getCell("G4").style = {
			font: fontBold,
			alignment: textLeft,
			border: borderThin,
		};
		worksheet.getCell("G5").style = {
			font: fontBold,
			alignment: textLeft,
			border: borderThin,
		};
		worksheet.getCell("G6").style = {
			font: fontBold,
			alignment: textLeft,
			border: borderThin,
		};
		worksheet.getCell("G7").style = {
			font: fontBold,
			alignment: textLeft,
			border: borderThin,
		};
		worksheet.getCell("G8").style = {
			font: fontBold,
			alignment: textLeft,
			border: borderThin,
		};
		worksheet.getCell("G9").style = {
			font: fontBold,
			alignment: textLeft,
			border: borderThin,
		};
		worksheet.getCell("G10").style = {
			font: fontBold,
			alignment: textLeft,
			border: borderThin,
		};
		worksheet.getCell("J4").style = {
			font: fontBold,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("K4").style = {
			font: fontBold,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("L4").style = {
			font: fontBold,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("M4").style = {
			font: fontBold,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("N4").style = {
			font: fontBold,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("J5").style = {
			font: fontBold,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("J7").style = {
			font: fontBold,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("J9").style = {
			font: fontBold,
			alignment: textCenter,
			border: borderThin,
		};

		worksheet.getCell("O4").style = {
			font: fontBold,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("O5").style = {
			font: fontBold,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("O7").style = {
			font: fontBold,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("O9").style = {
			font: fontBold,
			alignment: textCenter,
			border: borderThin,
			fill: yellowHeader,
		};
		worksheet.getCell("P5").style = {
			font: fontBold,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("Q5").style = {
			font: fontBold,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("R5").style = {
			font: fontBold,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("S4").style = {
			font: fontBold,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("S5").style = {
			font: fontBold,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("T5").style = {
			font: fontBold,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("U5").style = {
			font: fontBold,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("V5").style = {
			font: fontBold,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("W4").style = {
			fill: yellowHeader,
			font: fontBold,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("W5").style = {
			fill: yellowHeader,
			font: fontBold,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("X5").style = {
			fill: yellowHeader,
			font: fontBold,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("Y5").style = {
			fill: yellowHeader,
			font: fontBold,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("Z5").style = {
			fill: yellowHeader,
			font: fontBold,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("A11:K11").style = {
			font: fontBold,
			alignment: textCenter,
			fill: grayHeader,
			border: borderThin,
		};
		worksheet.getCell("L11").style = {
			font: fontBold,
			alignment: textCenter,
			fill: grayHeader,
			border: borderThin,
		};
		worksheet.getCell("A12").style = {
			font: fontBold,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("B12").style = {
			font: fontBold,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("D12").style = {
			font: fontBold,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("F12").style = {
			font: fontBold,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("H12").style = {
			font: fontBold,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("J12").style = {
			font: fontBold,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("B13").style = {
			font: fontBold,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("C13").style = {
			font: fontBold,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("D13").style = {
			font: fontBold,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("E13").style = {
			font: fontBold,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("F13").style = {
			font: fontBold,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("G13").style = {
			font: fontBold,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("H13").style = {
			font: fontBold,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("I13").style = {
			font: fontBold,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("J13").style = {
			font: fontBold,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("K13").style = {
			font: fontBold,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("L12").style = {
			font: fontBold,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("M12").style = {
			font: fontBold,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("O12").style = {
			font: fontBold,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("Q12").style = {
			font: fontBold,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("S12").style = {
			font: fontBold,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("U12").style = {
			font: fontBold,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("O13").style = {
			font: fontBold,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("P13").style = {
			font: fontBold,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("Q13").style = {
			font: fontBold,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("R13").style = {
			font: fontBold,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("S13").style = {
			font: fontBold,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("T13").style = {
			font: fontBold,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("U13").style = {
			font: fontBold,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("V13").style = {
			font: fontBold,
			alignment: textCenter,
			border: borderThin,
		};

		//column for data
		worksheet.getCell("B4").style = {
			border: borderThin,
			alignment: textCenter,
		};
		worksheet.getCell("B6").style = {
			border: borderThin,
			alignment: textCenter,
		};
		worksheet.getCell("B8").style = {
			border: borderThin,
			alignment: textCenter,
		};
		worksheet.getCell("E4").style = {
			border: borderThin,
			alignment: textCenter,
		};
		worksheet.getCell("E5").style = {
			border: borderThin,
			alignment: textCenter,
		};
		worksheet.getCell("E6").style = {
			border: borderThin,
			alignment: textCenter,
		};
		worksheet.getCell("E7").style = {
			border: borderThin,
			alignment: textCenter,
		};
		worksheet.getCell("E8").style = {
			border: borderThin,
			alignment: textCenter,
		};
		worksheet.getCell("E9").style = {
			border: borderThin,
			alignment: textCenter,
		};
		worksheet.getCell("E10").style = {
			border: borderThin,
			alignment: textCenter,
		};
		worksheet.getCell("H4").style = {
			border: borderThin,
			alignment: textCenter,
		};
		worksheet.getCell("H5").style = {
			border: borderThin,
			alignment: textCenter,
		};
		worksheet.getCell("H6").style = {
			border: borderThin,
			alignment: textCenter,
		};
		worksheet.getCell("H7").style = {
			border: borderThin,
			alignment: textCenter,
		};
		worksheet.getCell("H8").style = {
			border: borderThin,
			alignment: textCenter,
		};
		worksheet.getCell("H9").style = {
			border: borderThin,
			alignment: textCenter,
		};
		worksheet.getCell("H10").style = {
			border: borderThin,
			alignment: textCenter,
		};
		worksheet.getCell("K5").style = {
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("K7").style = {
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("L5").style = {
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("L7").style = {
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("M5").style = {
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("M7").style = {
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("N5").style = {
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("N7").style = {
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("N9").style = {
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("O6").style = {
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("P6").style = {
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("Q4").style = {
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("Q6").style = {
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("R6").style = {
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("S6").style = {
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("T6").style = {
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("U4").style = {
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("U6").style = {
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("V6").style = {
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("S7").style = {
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("S9").style = {
			fill: yellowHeader,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("W6").style = {
			fill: yellowHeader,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("X6").style = {
			fill: yellowHeader,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("Y4").style = {
			fill: yellowHeader,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("Y6").style = {
			fill: yellowHeader,
			alignment: textCenter,
			border: borderThin,
		};
		worksheet.getCell("Z6").style = {
			fill: yellowHeader,
			alignment: textCenter,
			border: borderThin,
		};

		let datePrint = `${moment().format("YYYYMMDD")}_${moment().format(
			"HHmmss"
		)}`;
		await workbook.xlsx
			.writeFile(
				`${path}/COAL_TERMINAL_SUMMARY_REPORT_12_HOURS_SHIFT_PAGE_1_${datePrint}.xlsx`
			)
			.then(() => {
				res.download(
					`${path}/COAL_TERMINAL_SUMMARY_REPORT_12_HOURS_SHIFT_PAGE_1_${datePrint}.xlsx`,
					`COAL_TERMINAL_SUMMARY_REPORT_12_HOURS_SHIFT_PAGE_2_${datePrint}.xlsx`,
					(err) => {
						if (err) {
							console.log(err);
						} else {
							fs.unlinkSync(
								`${path}/COAL_TERMINAL_SUMMARY_REPORT_12_HOURS_SHIFT_PAGE_2_${datePrint}.xlsx`
							);
						}
					}
				);
			});
	} catch (error) {
		console.log(error);
		res.json({ status: "failed", reason: error });
	}
};

const exportExcelPage2 = async (req, res) => {
	try {
		const wb = new ExcelJS.Workbook();
		const ws = wb.addWorksheet("Page 2", {
			properties: { tabColor: { argb: "FFC0000" } },
		});
		// const path = "./public";

		ws.mergeCells("A1:U1");
		ws.getCell("A1").value = "COAL TERMINAL SUMMARY REPORT 12 HOURS SHIFT";
		ws.getCell("A1").font = fontBolder;
		ws.getCell("A1").alignment = textCenter;

		ws.mergeCells("O2:Q2");
		ws.getCell("O2").value = "DATE";
		ws.getCell("O2").font = fontBold;
		ws.getCell("O2").alignment = textCenter;
		ws.getCell("O2").border = borderBold;
		ws.getCell("O2").fill = grayHeader;

		ws.mergeCells("O3:Q3");
		ws.getCell("O3").value = moment().format("YYYY/MM/DD HH:mm:ss");
		ws.getCell("O3").alignment = textCenter;
		ws.getCell("O3").border = borderThin;

		ws.mergeCells("R2:T2");
		ws.getCell("R2").value = "SHIFT";
		ws.getCell("R2").font = fontBold;
		ws.getCell("R2").alignment = textCenter;
		ws.getCell("R2").border = borderBold;
		ws.getCell("R2").fill = grayHeader;

		ws.mergeCells("R3:T3");
		ws.getCell("R3").value = "MALAM";
		ws.getCell("R3").alignment = textCenter;
		ws.getCell("R3").border = borderThin;

		ws.mergeCells("U2:W2");
		ws.getCell("U2").value = "CREW";
		ws.getCell("U2").font = fontBold;
		ws.getCell("U2").alignment = textCenter;
		ws.getCell("U2").border = borderBold;
		ws.getCell("U2").fill = grayHeader;

		ws.mergeCells("U3:W3");
		ws.getCell("U3").value = "PAK ANAS";
		ws.getCell("U3").alignment = textCenter;
		ws.getCell("U3").border = borderThin;

		ws.mergeCells("A4:W4");
		ws.getCell("A4").value = "DELAY DETAIL";
		ws.getCell("A4").font = fontBold;
		ws.getCell("A4").alignment = textCenter;
		ws.getCell("A4").border = borderBold;
		ws.getCell("A4").fill = grayHeader;

		ws.mergeCells("A5:B5");
		ws.getCell("A5").value = "Stream";
		ws.getCell("A5").font = fontBold;
		ws.getCell("A5").alignment = textCenter;
		ws.getCell("A5").border = borderThin;

		ws.mergeCells("C5:D5");
		ws.getCell("C5").value = "Start";
		ws.getCell("C5").font = fontBold;
		ws.getCell("C5").alignment = textCenter;
		ws.getCell("C5").border = borderThin;

		ws.mergeCells("E5:F5");
		ws.getCell("E5").value = "Stop";
		ws.getCell("E5").font = fontBold;
		ws.getCell("E5").alignment = textCenter;
		ws.getCell("E5").border = borderThin;

		ws.mergeCells("G5:H5");
		ws.getCell("G5").value = "Delay Time";
		ws.getCell("G5").font = fontBold;
		ws.getCell("G5").alignment = textCenter;
		ws.getCell("G5").border = borderThin;

		ws.mergeCells("I5:J5");
		ws.getCell("I5").value = "Delay Cat";
		ws.getCell("I5").font = fontBold;
		ws.getCell("I5").alignment = textCenter;
		ws.getCell("I5").border = borderThin;

		ws.mergeCells("K5:L5");
		ws.getCell("K5").value = "Equipment";
		ws.getCell("K5").font = fontBold;
		ws.getCell("K5").alignment = textCenter;
		ws.getCell("K5").border = borderThin;

		ws.mergeCells("M5:N5");
		ws.getCell("M5").value = "Delay Code";
		ws.getCell("M5").font = fontBold;
		ws.getCell("M5").alignment = textCenter;
		ws.getCell("M5").border = borderThin;

		ws.mergeCells("O5:W5");
		ws.getCell("O5").value = "Comment";
		ws.getCell("O5").font = fontBold;
		ws.getCell("O5").alignment = textCenter;
		ws.getCell("O5").border = borderThin;

		let data = [
			{
				straem: "straem 1",
				start: "2024/05/08 10:00:00",
				stop: "2024/05/08 19:00:00",
				delay_time: "9 hours",
				delay_cat: "unknown",
				delay_coce: "unknown_code",
				comment: "no comment",
			},
			{
				straem: "straem 1",
				start: "2024/05/08 10:00:00",
				stop: "2024/05/08 19:00:00",
				delay_time: "9 hours",
				delay_cat: "unknown",
				delay_coce: "unknown_code",
				comment: "no comment",
			},
			{
				straem: "straem 1",
				start: "2024/05/08 10:00:00",
				stop: "2024/05/08 19:00:00",
				delay_time: "9 hours",
				delay_cat: "unknown",
				delay_coce: "unknown_code",
				comment: "no comment",
			},
			{
				straem: "straem 1",
				start: "2024/05/08 10:00:00",
				stop: "2024/05/08 19:00:00",
				delay_time: "9 hours",
				delay_cat: "unknown",
				delay_coce: "unknown_code",
				comment: "no comment",
			},
			{
				straem: "straem 1",
				start: "2024/05/08 10:00:00",
				stop: "2024/05/08 19:00:00",
				delay_time: "9 hours",
				delay_cat: "unknown",
				delay_coce: "unknown_code",
				comment: "no comment",
			},
			{
				straem: "straem 1",
				start: "2024/05/08 10:00:00",
				stop: "2024/05/08 19:00:00",
				delay_time: "9 hours",
				delay_cat: "unknown",
				delay_coce: "unknown_code",
				comment: "no comment",
			},
			{
				straem: "straem 1",
				start: "2024/05/08 10:00:00",
				stop: "2024/05/08 19:00:00",
				delay_time: "9 hours",
				delay_cat: "unknown",
				equipment: "CRUSHER",
				delay_code: "unknown_code",
				comment: "no comment",
			},
		];

		let startData = 5;

		for (let i in data) {
			startData++;
			ws.mergeCells("A" + startData + ":" + "B" + startData);
			ws.getCell("A" + startData).value = data[i].straem;
			ws.getCell("A" + startData).alignment = textCenter;
			ws.getCell("A" + startData).border = borderThin;

			ws.mergeCells("C" + startData + ":" + "D" + startData);
			ws.getCell("C" + startData).value = data[i].start;
			ws.getCell("C" + startData).alignment = textCenter;
			ws.getCell("C" + startData).border = borderThin;

			ws.mergeCells("E" + startData + ":" + "F" + startData);
			ws.getCell("E" + startData).value = data[i].stop;
			ws.getCell("E" + startData).alignment = textCenter;
			ws.getCell("E" + startData).border = borderThin;

			ws.mergeCells("G" + startData + ":" + "H" + startData);
			ws.getCell("G" + startData).value = data[i].delay_time;
			ws.getCell("G" + startData).alignment = textCenter;
			ws.getCell("G" + startData).border = borderThin;

			ws.mergeCells("I" + startData + ":" + "J" + startData);
			ws.getCell("I" + startData).value = data[i].delay_cat;
			ws.getCell("I" + startData).alignment = textCenter;
			ws.getCell("I" + startData).border = borderThin;

			ws.mergeCells("K" + startData + ":" + "L" + startData);
			ws.getCell("K" + startData).value = data[i].equipment;
			ws.getCell("K" + startData).alignment = textCenter;
			ws.getCell("K" + startData).border = borderThin;

			ws.mergeCells("M" + startData + ":" + "N" + startData);
			ws.getCell("M" + startData).value = data[i].delay_code;
			ws.getCell("M" + startData).alignment = textCenter;
			ws.getCell("M" + startData).border = borderThin;

			ws.mergeCells("O" + startData + ":" + "W" + startData);
			ws.getCell("O" + startData).value = data[i].comment;
			ws.getCell("O" + startData).alignment = textCenter;
			ws.getCell("O" + startData).border = borderThin;
		}

		let datePrint = `${moment().format("YYYYMMDD")}_${moment().format(
			"HHmmss"
		)}`;
		const direct = path.join(process.cwd(), "tmp");
		await wb.xlsx
			.writeFile(
				`${direct}/COAL_TERMINAL_SUMMARY_REPORT_12_HOURS_SHIFT_PAGE_2_${datePrint}.xlsx`
			)
			.then(() => {
				console.log("HARUS NYA SIH UDA BISA DOWNLOAD YA");
				res.download(
					`${direct}/COAL_TERMINAL_SUMMARY_REPORT_12_HOURS_SHIFT_PAGE_2_${datePrint}.xlsx`,
					`COAL_TERMINAL_SUMMARY_REPORT_12_HOURS_SHIFT_PAGE_2_${datePrint}.xlsx`,
					(err) => {
						if (err) {
							console.log(err);
						} else {
							console.log("err");
							fs.unlinkSync(
								`${direct}/COAL_TERMINAL_SUMMARY_REPORT_12_HOURS_SHIFT_PAGE_2_${datePrint}.xlsx`
							);
						}
					}
				);
				// res.json({ message: "success" });
			});
	} catch (error) {
		res.json({ status: "failed", reason: error });
	}
};

module.exports = {
	exportExcel,
	exportExcelPage2,
};
