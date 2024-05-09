const ExcelJS = require("exceljs");
const fs = require("fs");
const {
	fontBold,
	textCenter,
	cocoHeader,
	borderBold,
	borderThin,
	yellowHeader,
	textRight,
} = require("../utils/excel");
const moment = require("moment");

const exportExcel = async (req, res) => {
	try {
		const wb = new ExcelJS.Workbook();
		const ws = wb.addWorksheet("New Sheet", {
			properties: { tabColor: { argb: "FFC0000" } },
		});

		// tail olc 1 header
		ws.mergeCells("B5:K5");
		ws.getCell("B5").value = "TAILEND WEIGHER (OLC-1)";
		ws.getCell("B5").font = fontBold;
		ws.getCell("B5").alignment = textCenter;
		ws.getCell("B5").border = borderThin;
		ws.getCell("B5").fill = cocoHeader;

		// tail olc 2 header
		ws.mergeCells("L5:U5");
		ws.getCell("L5").value = "TAILEND WEIGHER (OLC-2)";
		ws.getCell("L5").font = fontBold;
		ws.getCell("L5").alignment = textCenter;
		ws.getCell("L5").border = borderThin;
		ws.getCell("L5").fill = cocoHeader;

		//tail olc 1 prima header
		ws.mergeCells("B6:C6");
		ws.getCell("B6").value = "PRIMA";
		ws.getCell("B6").font = fontBold;
		ws.getCell("B6").alignment = textCenter;
		ws.getCell("B6").border = borderThin;
		ws.getCell("B6").fill = cocoHeader;
		//tail olc 1 pinang header
		ws.mergeCells("D6:E6");
		ws.getCell("D6").value = "PINANG";
		ws.getCell("D6").font = fontBold;
		ws.getCell("D6").alignment = textCenter;
		ws.getCell("D6").border = borderThin;
		ws.getCell("D6").fill = cocoHeader;
		//tail olc 1 melawan header
		ws.mergeCells("F6:G6");
		ws.getCell("F6").value = "MELAWAN";
		ws.getCell("F6").font = fontBold;
		ws.getCell("F6").alignment = textCenter;
		ws.getCell("F6").border = borderThin;
		ws.getCell("F6").fill = cocoHeader;
		//tail olc 1 mlw42 header
		ws.mergeCells("H6:I6");
		ws.getCell("H6").value = "MLW42";
		ws.getCell("H6").font = fontBold;
		ws.getCell("H6").alignment = textCenter;
		ws.getCell("H6").border = borderThin;
		ws.getCell("H6").fill = cocoHeader;
		//tail olc 1 total header
		ws.mergeCells("J6:K6");
		ws.getCell("J6").value = "TOTAL";
		ws.getCell("J6").font = fontBold;
		ws.getCell("J6").alignment = textCenter;
		ws.getCell("J6").border = borderThin;
		ws.getCell("J6").fill = cocoHeader;

		//tail olc 2 prima header
		ws.mergeCells("L6:M6");
		ws.getCell("L6").value = "PRIMA";
		ws.getCell("L6").font = fontBold;
		ws.getCell("L6").alignment = textCenter;
		ws.getCell("L6").border = borderThin;
		ws.getCell("L6").fill = cocoHeader;
		//tail olc 2 pinang header
		ws.mergeCells("N6:O6");
		ws.getCell("N6").value = "PINANG";
		ws.getCell("N6").font = fontBold;
		ws.getCell("N6").alignment = textCenter;
		ws.getCell("N6").border = borderThin;
		ws.getCell("N6").fill = cocoHeader;
		//tail olc 2 melawan header
		ws.mergeCells("P6:Q6");
		ws.getCell("P6").value = "MELAWAN";
		ws.getCell("P6").font = fontBold;
		ws.getCell("P6").alignment = textCenter;
		ws.getCell("P6").border = borderThin;
		ws.getCell("P6").fill = cocoHeader;
		//tail olc 2 mlw42 header
		ws.mergeCells("R6:S6");
		ws.getCell("R6").value = "MLW42";
		ws.getCell("R6").font = fontBold;
		ws.getCell("R6").alignment = textCenter;
		ws.getCell("R6").border = borderThin;
		ws.getCell("R6").fill = cocoHeader;
		//tail olc 2 total header
		ws.mergeCells("T6:U6");
		ws.getCell("T6").value = "TOTAL";
		ws.getCell("T6").font = fontBold;
		ws.getCell("T6").alignment = textCenter;
		ws.getCell("T6").border = borderThin;
		ws.getCell("T6").fill = cocoHeader;

		// TOTAL CONVIYED HEADER
		ws.mergeCells("V5:V6");
		ws.getCell("V6").value = "TOTAL COAL CONVEYED (AT TAIL END)";
		ws.getCell("V6").font = fontBold;
		ws.getCell("V6").alignment = textCenter;
		ws.getCell("V6").border = borderThin;
		ws.getCell("V6").fill = cocoHeader;

		let dataTailEnd = [
			{
				b: "prima",
				d: "pinang",
				f: "melawan",
				h: "mlw42",
				j: "totalnya",
				l: "prima",
				n: "pinang",
				p: "melawan",
				r: "mlw42",
				t: "totalnya",
				v: "conveyed",
			},
		];
		let startTail = 6;

		for (let i in dataTailEnd) {
			startTail++;
			// tail olc 1 header
			ws.mergeCells("B" + startTail + ":" + "C" + startTail);
			ws.getCell("B" + startTail).value = dataTailEnd[i].b;
			ws.getCell("B" + startTail).alignment = textCenter;
			ws.getCell("B" + startTail).border = borderThin;

			// tail olc 2 header
			ws.mergeCells("D" + startTail + ":" + "E" + startTail);
			ws.getCell("D" + startTail).value = dataTailEnd[i].d;
			ws.getCell("D" + startTail).alignment = textCenter;
			ws.getCell("D" + startTail).border = borderThin;

			//tail olc 1 prima header
			ws.mergeCells("F" + startTail + ":" + "G" + startTail);
			ws.getCell("F" + startTail).value = dataTailEnd[i].f;
			ws.getCell("F" + startTail).alignment = textCenter;
			ws.getCell("F" + startTail).border = borderThin;

			//tail olc 1 pinang header
			ws.mergeCells("H" + startTail + ":" + "I" + startTail);
			ws.getCell("H" + startTail).value = dataTailEnd[i].h;
			ws.getCell("H" + startTail).alignment = textCenter;
			ws.getCell("H" + startTail).border = borderThin;

			//tail olc 1 melawan header
			ws.mergeCells("J" + startTail + ":" + "K" + startTail);
			ws.getCell("J" + startTail).value = dataTailEnd[i].j;
			ws.getCell("J" + startTail).alignment = textCenter;
			ws.getCell("J" + startTail).border = borderThin;

			//tail olc 1 mlw42 header
			ws.mergeCells("L" + startTail + ":" + "M" + startTail);
			ws.getCell("L" + startTail).value = dataTailEnd[i].l;
			ws.getCell("L" + startTail).alignment = textCenter;
			ws.getCell("L" + startTail).border = borderThin;

			//tail olc 1 total header
			ws.mergeCells("N" + startTail + ":" + "O" + startTail);
			ws.getCell("N" + startTail).value = dataTailEnd[i].n;
			ws.getCell("N" + startTail).alignment = textCenter;
			ws.getCell("N" + startTail).border = borderThin;

			//tail olc 2 prima header
			ws.mergeCells("P" + startTail + ":" + "Q" + startTail);
			ws.getCell("P" + startTail).value = dataTailEnd[i].p;
			ws.getCell("P" + startTail).alignment = textCenter;
			ws.getCell("P" + startTail).border = borderThin;

			//tail olc 2 pinang header
			ws.mergeCells("R" + startTail + ":" + "S" + startTail);
			ws.getCell("R" + startTail).value = dataTailEnd[i].r;
			ws.getCell("R" + startTail).alignment = textCenter;
			ws.getCell("R" + startTail).border = borderThin;

			//tail olc 2 melawan header
			ws.mergeCells("T" + startTail + ":" + "U" + startTail);
			ws.getCell("T" + startTail).value = dataTailEnd[i].t;
			ws.getCell("T" + startTail).alignment = textCenter;
			ws.getCell("T" + startTail).border = borderThin;

			//tail olc 2 mlw42 header
			ws.getCell("V" + startTail).value = dataTailEnd[i].v;
			ws.getCell("V" + startTail).alignment = textCenter;
			ws.getCell("V" + startTail).border = borderThin;
		}
		//Head end olc 1 header
		let startHeadEnd = startTail + 1;
		let secondHeadEnd = startHeadEnd + 1;
		// Head end olc 1 header
		ws.mergeCells("B" + startHeadEnd + ":" + "K" + startHeadEnd);
		ws.getCell("B" + startHeadEnd).value = "HEAD END VIRTUAL WEIGHER (OLC-1)";
		ws.getCell("B" + startHeadEnd).font = fontBold;
		ws.getCell("B" + startHeadEnd).alignment = textCenter;
		ws.getCell("B" + startHeadEnd).border = borderThin;
		ws.getCell("B" + startHeadEnd).fill = cocoHeader;

		// Head end olc 2 header
		ws.mergeCells("L" + startHeadEnd + ":" + "U" + startHeadEnd);
		ws.getCell("L" + startHeadEnd).value =
			"RECLAIMING-2 & STACKING-2 CONVEYOR WEIGHER";
		ws.getCell("L" + startHeadEnd).font = fontBold;
		ws.getCell("L" + startHeadEnd).alignment = textCenter;
		ws.getCell("L" + startHeadEnd).border = borderThin;
		ws.getCell("L" + startHeadEnd).fill = cocoHeader;

		//Head end olc 1 prima header
		ws.mergeCells("B" + secondHeadEnd + ":" + "C" + secondHeadEnd);
		ws.getCell("B" + secondHeadEnd).value = "PRIMA";
		ws.getCell("B" + secondHeadEnd).font = fontBold;
		ws.getCell("B" + secondHeadEnd).alignment = textCenter;
		ws.getCell("B" + secondHeadEnd).border = borderThin;
		ws.getCell("B" + secondHeadEnd).fill = cocoHeader;
		//Head end olc 1 pinang header
		ws.mergeCells("D" + secondHeadEnd + ":" + "E" + secondHeadEnd);
		ws.getCell("D" + secondHeadEnd).value = "PINANG";
		ws.getCell("D" + secondHeadEnd).font = fontBold;
		ws.getCell("D" + secondHeadEnd).alignment = textCenter;
		ws.getCell("D" + secondHeadEnd).border = borderThin;
		ws.getCell("D" + secondHeadEnd).fill = cocoHeader;
		//Head end olc 1 melawan header
		ws.mergeCells("F" + secondHeadEnd + ":" + "G" + secondHeadEnd);
		ws.getCell("F" + secondHeadEnd).value = "MELAWAN";
		ws.getCell("F" + secondHeadEnd).font = fontBold;
		ws.getCell("F" + secondHeadEnd).alignment = textCenter;
		ws.getCell("F" + secondHeadEnd).border = borderThin;
		ws.getCell("F" + secondHeadEnd).fill = cocoHeader;
		//Head end olc 1 mlw42 header
		ws.mergeCells("H" + secondHeadEnd + ":" + "I" + secondHeadEnd);
		ws.getCell("H" + secondHeadEnd).value = "MLW42";
		ws.getCell("H" + secondHeadEnd).font = fontBold;
		ws.getCell("H" + secondHeadEnd).alignment = textCenter;
		ws.getCell("H" + secondHeadEnd).border = borderThin;
		ws.getCell("H" + secondHeadEnd).fill = cocoHeader;
		//Head end olc 1 total header
		ws.mergeCells("J" + secondHeadEnd + ":" + "K" + secondHeadEnd);
		ws.getCell("J" + secondHeadEnd).value = "TOTAL";
		ws.getCell("J" + secondHeadEnd).font = fontBold;
		ws.getCell("J" + secondHeadEnd).alignment = textCenter;
		ws.getCell("J" + secondHeadEnd).border = borderThin;
		ws.getCell("J" + secondHeadEnd).fill = cocoHeader;

		//Head end olc 2 prima header
		ws.mergeCells("L" + secondHeadEnd + ":" + "M" + secondHeadEnd);
		ws.getCell("L" + secondHeadEnd).value = "PRIMA";
		ws.getCell("L" + secondHeadEnd).font = fontBold;
		ws.getCell("L" + secondHeadEnd).alignment = textCenter;
		ws.getCell("L" + secondHeadEnd).border = borderThin;
		ws.getCell("L" + secondHeadEnd).fill = cocoHeader;
		//Head end olc 2 pinang header
		ws.mergeCells("N" + secondHeadEnd + ":" + "O" + secondHeadEnd);
		ws.getCell("N" + secondHeadEnd).value = "PINANG";
		ws.getCell("N" + secondHeadEnd).font = fontBold;
		ws.getCell("N" + secondHeadEnd).alignment = textCenter;
		ws.getCell("N" + secondHeadEnd).border = borderThin;
		ws.getCell("N" + secondHeadEnd).fill = cocoHeader;
		//Head end olc 2 melawan header
		ws.mergeCells("P" + secondHeadEnd + ":" + "Q" + secondHeadEnd);
		ws.getCell("P" + secondHeadEnd).value = "MELAWAN";
		ws.getCell("P" + secondHeadEnd).font = fontBold;
		ws.getCell("P" + secondHeadEnd).alignment = textCenter;
		ws.getCell("P" + secondHeadEnd).border = borderThin;
		ws.getCell("P" + secondHeadEnd).fill = cocoHeader;
		//Head end olc 2 mlw42 header
		ws.mergeCells("R" + secondHeadEnd + ":" + "S" + secondHeadEnd);
		ws.getCell("R" + secondHeadEnd).value = "MLW42";
		ws.getCell("R" + secondHeadEnd).font = fontBold;
		ws.getCell("R" + secondHeadEnd).alignment = textCenter;
		ws.getCell("R" + secondHeadEnd).border = borderThin;
		ws.getCell("R" + secondHeadEnd).fill = cocoHeader;
		//Head end olc 2 total header
		ws.mergeCells("T" + secondHeadEnd + ":" + "U" + secondHeadEnd);
		ws.getCell("T" + secondHeadEnd).value = "TOTAL";
		ws.getCell("T" + secondHeadEnd).font = fontBold;
		ws.getCell("T" + secondHeadEnd).alignment = textCenter;
		ws.getCell("T" + secondHeadEnd).border = borderThin;
		ws.getCell("T" + secondHeadEnd).fill = cocoHeader;

		// TOTAL CONVIYED HEADER
		ws.mergeCells("V" + startHeadEnd + ":" + "V" + secondHeadEnd);
		ws.getCell("V" + startHeadEnd).value = "TOTAL COAL CONVEYED (AT TBCT)";
		ws.getCell("V" + startHeadEnd).font = fontBold;
		ws.getCell("V" + startHeadEnd).alignment = textCenter;
		ws.getCell("V" + startHeadEnd).border = borderThin;
		ws.getCell("V" + startHeadEnd).fill = cocoHeader;

		let dataHeadEnd = [
			{
				b: "prima",
				d: "pinang",
				f: "melawan",
				h: "mlw42",
				j: "totalnya",
				l: "prima",
				n: "pinang",
				p: "melawan",
				r: "mlw42",
				t: "totalnya",
				v: "conveyed",
			},
		];

		let startHead = secondHeadEnd;

		for (let i in dataHeadEnd) {
			startHead++;
			// tail olc 1 header
			ws.mergeCells("B" + startHead + ":" + "C" + startHead);
			ws.getCell("B" + startHead).value = dataHeadEnd[i].b;
			ws.getCell("B" + startHead).alignment = textCenter;
			ws.getCell("B" + startHead).border = borderThin;

			// tail olc 2 header
			ws.mergeCells("D" + startHead + ":" + "E" + startHead);
			ws.getCell("D" + startHead).value = dataHeadEnd[i].d;
			ws.getCell("D" + startHead).alignment = textCenter;
			ws.getCell("D" + startHead).border = borderThin;

			//tail olc 1 prima header
			ws.mergeCells("F" + startHead + ":" + "G" + startHead);
			ws.getCell("F" + startHead).value = dataHeadEnd[i].f;
			ws.getCell("F" + startHead).alignment = textCenter;
			ws.getCell("F" + startHead).border = borderThin;

			//tail olc 1 pinang header
			ws.mergeCells("H" + startHead + ":" + "I" + startHead);
			ws.getCell("H" + startHead).value = dataHeadEnd[i].h;
			ws.getCell("H" + startHead).alignment = textCenter;
			ws.getCell("H" + startHead).border = borderThin;

			//tail olc 1 melawan header
			ws.mergeCells("J" + startHead + ":" + "K" + startHead);
			ws.getCell("J" + startHead).value = dataHeadEnd[i].j;
			ws.getCell("J" + startHead).alignment = textCenter;
			ws.getCell("J" + startHead).border = borderThin;

			//tail olc 1 mlw42 header
			ws.mergeCells("L" + startHead + ":" + "M" + startHead);
			ws.getCell("L" + startHead).value = dataHeadEnd[i].l;
			ws.getCell("L" + startHead).alignment = textCenter;
			ws.getCell("L" + startHead).border = borderThin;

			//tail olc 1 total header
			ws.mergeCells("N" + startHead + ":" + "O" + startHead);
			ws.getCell("N" + startHead).value = dataHeadEnd[i].n;
			ws.getCell("N" + startHead).alignment = textCenter;
			ws.getCell("N" + startHead).border = borderThin;

			//tail olc 2 prima header
			ws.mergeCells("P" + startHead + ":" + "Q" + startHead);
			ws.getCell("P" + startHead).value = dataHeadEnd[i].p;
			ws.getCell("P" + startHead).alignment = textCenter;
			ws.getCell("P" + startHead).border = borderThin;

			//tail olc 2 pinang header
			ws.mergeCells("R" + startHead + ":" + "S" + startHead);
			ws.getCell("R" + startHead).value = dataHeadEnd[i].r;
			ws.getCell("R" + startHead).alignment = textCenter;
			ws.getCell("R" + startHead).border = borderThin;

			//tail olc 2 melawan header
			ws.mergeCells("T" + startHead + ":" + "U" + startHead);
			ws.getCell("T" + startHead).value = dataHeadEnd[i].t;
			ws.getCell("T" + startHead).alignment = textCenter;
			ws.getCell("T" + startHead).border = borderThin;

			//tail olc 2 mlw42 header
			ws.getCell("V" + startHead).value = dataHeadEnd[i].v;
			ws.getCell("V" + startHead).alignment = textCenter;
			ws.getCell("V" + startHead).border = borderThin;
		}

		let startCoal = startHead + 2;
		// coal header
		ws.mergeCells("B" + startCoal + ":" + "K" + startCoal);
		ws.getCell("B" + startCoal).value = "COAL STACKING (INCOMING)";
		ws.getCell("B" + startCoal).font = fontBold;
		ws.getCell("B" + startCoal).alignment = textCenter;
		ws.getCell("B" + startCoal).border = borderThin;
		ws.getCell("B" + startCoal).fill = cocoHeader;

		ws.mergeCells("B" + (startCoal + 1) + ":" + "I" + (startCoal + 1));
		ws.getCell("B" + (startCoal + 1)).value = "STACKER-1";
		ws.getCell("B" + (startCoal + 1)).font = fontBold;
		ws.getCell("B" + (startCoal + 1)).alignment = textCenter;
		ws.getCell("B" + (startCoal + 1)).border = borderThin;
		ws.getCell("B" + (startCoal + 1)).fill = cocoHeader;

		ws.mergeCells("J" + (startCoal + 1) + ":" + "K" + (startCoal + 2));
		ws.getCell("J" + (startCoal + 1)).value = "TOTAL";
		ws.getCell("J" + (startCoal + 1)).font = fontBold;
		ws.getCell("J" + (startCoal + 1)).alignment = textCenter;
		ws.getCell("J" + (startCoal + 1)).border = borderThin;
		ws.getCell("J" + (startCoal + 1)).fill = cocoHeader;

		ws.mergeCells("B" + (startCoal + 2) + ":" + "C" + (startCoal + 2));
		ws.getCell("B" + (startCoal + 2)).value = "PRIMA";
		ws.getCell("B" + (startCoal + 2)).font = fontBold;
		ws.getCell("B" + (startCoal + 2)).alignment = textCenter;
		ws.getCell("B" + (startCoal + 2)).border = borderThin;
		ws.getCell("B" + (startCoal + 2)).fill = cocoHeader;
		//Head end olc 1 pinang header
		ws.mergeCells("D" + (startCoal + 2) + ":" + "E" + (startCoal + 2));
		ws.getCell("D" + (startCoal + 2)).value = "PINANG";
		ws.getCell("D" + (startCoal + 2)).font = fontBold;
		ws.getCell("D" + (startCoal + 2)).alignment = textCenter;
		ws.getCell("D" + (startCoal + 2)).border = borderThin;
		ws.getCell("D" + (startCoal + 2)).fill = cocoHeader;
		//Head end olc 1 melawan header
		ws.mergeCells("F" + (startCoal + 2) + ":" + "G" + (startCoal + 2));
		ws.getCell("F" + (startCoal + 2)).value = "MELAWAN";
		ws.getCell("F" + (startCoal + 2)).font = fontBold;
		ws.getCell("F" + (startCoal + 2)).alignment = textCenter;
		ws.getCell("F" + (startCoal + 2)).border = borderThin;
		ws.getCell("F" + (startCoal + 2)).fill = cocoHeader;
		//Head end olc 1 mlw42 header
		ws.mergeCells("H" + (startCoal + 2) + ":" + "I" + (startCoal + 2));
		ws.getCell("H" + (startCoal + 2)).value = "MLW42";
		ws.getCell("H" + (startCoal + 2)).font = fontBold;
		ws.getCell("H" + (startCoal + 2)).alignment = textCenter;
		ws.getCell("H" + (startCoal + 2)).border = borderThin;
		ws.getCell("H" + (startCoal + 2)).fill = cocoHeader;

		ws.getCell("B" + (startCoal + 3)).value = "SOUTH";
		ws.getCell("B" + (startCoal + 3)).font = fontBold;
		ws.getCell("B" + (startCoal + 3)).alignment = textCenter;
		ws.getCell("B" + (startCoal + 3)).border = borderThin;
		ws.getCell("B" + (startCoal + 3)).fill = cocoHeader;

		ws.getCell("C" + (startCoal + 3)).value = "NORTH";
		ws.getCell("C" + (startCoal + 3)).font = fontBold;
		ws.getCell("C" + (startCoal + 3)).alignment = textCenter;
		ws.getCell("C" + (startCoal + 3)).border = borderThin;
		ws.getCell("C" + (startCoal + 3)).fill = cocoHeader;

		ws.getCell("D" + (startCoal + 3)).value = "SOUTH";
		ws.getCell("D" + (startCoal + 3)).font = fontBold;
		ws.getCell("D" + (startCoal + 3)).alignment = textCenter;
		ws.getCell("D" + (startCoal + 3)).border = borderThin;
		ws.getCell("D" + (startCoal + 3)).fill = cocoHeader;

		ws.getCell("E" + (startCoal + 3)).value = "NORTH";
		ws.getCell("E" + (startCoal + 3)).font = fontBold;
		ws.getCell("E" + (startCoal + 3)).alignment = textCenter;
		ws.getCell("E" + (startCoal + 3)).border = borderThin;
		ws.getCell("E" + (startCoal + 3)).fill = cocoHeader;

		ws.getCell("F" + (startCoal + 3)).value = "SOUTH";
		ws.getCell("F" + (startCoal + 3)).font = fontBold;
		ws.getCell("F" + (startCoal + 3)).alignment = textCenter;
		ws.getCell("F" + (startCoal + 3)).border = borderThin;
		ws.getCell("F" + (startCoal + 3)).fill = cocoHeader;

		ws.getCell("G" + (startCoal + 3)).value = "NORTH";
		ws.getCell("G" + (startCoal + 3)).font = fontBold;
		ws.getCell("G" + (startCoal + 3)).alignment = textCenter;
		ws.getCell("G" + (startCoal + 3)).border = borderThin;
		ws.getCell("G" + (startCoal + 3)).fill = cocoHeader;

		ws.getCell("H" + (startCoal + 3)).value = "SOUTH";
		ws.getCell("H" + (startCoal + 3)).font = fontBold;
		ws.getCell("H" + (startCoal + 3)).alignment = textCenter;
		ws.getCell("H" + (startCoal + 3)).border = borderThin;
		ws.getCell("H" + (startCoal + 3)).fill = cocoHeader;

		ws.getCell("I" + (startCoal + 3)).value = "NORTH";
		ws.getCell("I" + (startCoal + 3)).font = fontBold;
		ws.getCell("I" + (startCoal + 3)).alignment = textCenter;
		ws.getCell("I" + (startCoal + 3)).border = borderThin;
		ws.getCell("I" + (startCoal + 3)).fill = cocoHeader;

		ws.getCell("J" + (startCoal + 3)).value = "SOUTH";
		ws.getCell("J" + (startCoal + 3)).font = fontBold;
		ws.getCell("J" + (startCoal + 3)).alignment = textCenter;
		ws.getCell("J" + (startCoal + 3)).border = borderThin;
		ws.getCell("J" + (startCoal + 3)).fill = cocoHeader;

		ws.getCell("K" + (startCoal + 3)).value = "NORTH";
		ws.getCell("K" + (startCoal + 3)).font = fontBold;
		ws.getCell("K" + (startCoal + 3)).alignment = textCenter;
		ws.getCell("K" + (startCoal + 3)).border = borderThin;
		ws.getCell("K" + (startCoal + 3)).fill = cocoHeader;

		// COAL RECLAIMING HEADER
		ws.mergeCells("L" + startCoal + ":" + "U" + startCoal);
		ws.getCell("L" + startCoal).value = "COAL RECLAIMING (OUTGOING)";
		ws.getCell("L" + startCoal).font = fontBold;
		ws.getCell("L" + startCoal).alignment = textCenter;
		ws.getCell("L" + startCoal).border = borderThin;
		ws.getCell("L" + startCoal).fill = yellowHeader;

		ws.mergeCells("L" + (startCoal + 1) + ":" + "S" + (startCoal + 1));
		ws.getCell("L" + (startCoal + 1)).value = "RECLAIMER-1";
		ws.getCell("L" + (startCoal + 1)).font = fontBold;
		ws.getCell("L" + (startCoal + 1)).alignment = textCenter;
		ws.getCell("L" + (startCoal + 1)).border = borderThin;
		ws.getCell("L" + (startCoal + 1)).fill = yellowHeader;

		ws.mergeCells("T" + (startCoal + 1) + ":" + "U" + (startCoal + 2));
		ws.getCell("T" + (startCoal + 1)).value = "TOTAL";
		ws.getCell("T" + (startCoal + 1)).font = fontBold;
		ws.getCell("T" + (startCoal + 1)).alignment = textCenter;
		ws.getCell("T" + (startCoal + 1)).border = borderThin;
		ws.getCell("T" + (startCoal + 1)).fill = yellowHeader;

		ws.mergeCells("L" + (startCoal + 2) + ":" + "M" + (startCoal + 2));
		ws.getCell("L" + (startCoal + 2)).value = "PRIMA";
		ws.getCell("L" + (startCoal + 2)).font = fontBold;
		ws.getCell("L" + (startCoal + 2)).alignment = textCenter;
		ws.getCell("L" + (startCoal + 2)).border = borderThin;
		ws.getCell("L" + (startCoal + 2)).fill = yellowHeader;
		//Head end olc 1 pinang header
		ws.mergeCells("N" + (startCoal + 2) + ":" + "O" + (startCoal + 2));
		ws.getCell("N" + (startCoal + 2)).value = "PINANG";
		ws.getCell("N" + (startCoal + 2)).font = fontBold;
		ws.getCell("N" + (startCoal + 2)).alignment = textCenter;
		ws.getCell("N" + (startCoal + 2)).border = borderThin;
		ws.getCell("N" + (startCoal + 2)).fill = yellowHeader;
		//Head end olc 1 melawan header
		ws.mergeCells("P" + (startCoal + 2) + ":" + "Q" + (startCoal + 2));
		ws.getCell("P" + (startCoal + 2)).value = "MELAWAN";
		ws.getCell("P" + (startCoal + 2)).font = fontBold;
		ws.getCell("P" + (startCoal + 2)).alignment = textCenter;
		ws.getCell("P" + (startCoal + 2)).border = borderThin;
		ws.getCell("P" + (startCoal + 2)).fill = yellowHeader;
		//Head end olc 1 mlw42 header
		ws.mergeCells("R" + (startCoal + 2) + ":" + "S" + (startCoal + 2));
		ws.getCell("R" + (startCoal + 2)).value = "MLW42";
		ws.getCell("R" + (startCoal + 2)).font = fontBold;
		ws.getCell("R" + (startCoal + 2)).alignment = textCenter;
		ws.getCell("R" + (startCoal + 2)).border = borderThin;
		ws.getCell("R" + (startCoal + 2)).fill = yellowHeader;

		ws.getCell("L" + (startCoal + 3)).value = "SOUTH";
		ws.getCell("L" + (startCoal + 3)).font = fontBold;
		ws.getCell("L" + (startCoal + 3)).alignment = textCenter;
		ws.getCell("L" + (startCoal + 3)).border = borderThin;
		ws.getCell("L" + (startCoal + 3)).fill = yellowHeader;

		ws.getCell("M" + (startCoal + 3)).value = "NORTH";
		ws.getCell("M" + (startCoal + 3)).font = fontBold;
		ws.getCell("M" + (startCoal + 3)).alignment = textCenter;
		ws.getCell("M" + (startCoal + 3)).border = borderThin;
		ws.getCell("M" + (startCoal + 3)).fill = yellowHeader;

		ws.getCell("N" + (startCoal + 3)).value = "SOUTH";
		ws.getCell("N" + (startCoal + 3)).font = fontBold;
		ws.getCell("N" + (startCoal + 3)).alignment = textCenter;
		ws.getCell("N" + (startCoal + 3)).border = borderThin;
		ws.getCell("N" + (startCoal + 3)).fill = yellowHeader;

		ws.getCell("O" + (startCoal + 3)).value = "NORTH";
		ws.getCell("O" + (startCoal + 3)).font = fontBold;
		ws.getCell("O" + (startCoal + 3)).alignment = textCenter;
		ws.getCell("O" + (startCoal + 3)).border = borderThin;
		ws.getCell("O" + (startCoal + 3)).fill = yellowHeader;

		ws.getCell("P" + (startCoal + 3)).value = "SOUTH";
		ws.getCell("P" + (startCoal + 3)).font = fontBold;
		ws.getCell("P" + (startCoal + 3)).alignment = textCenter;
		ws.getCell("P" + (startCoal + 3)).border = borderThin;
		ws.getCell("P" + (startCoal + 3)).fill = yellowHeader;

		ws.getCell("Q" + (startCoal + 3)).value = "NORTH";
		ws.getCell("Q" + (startCoal + 3)).font = fontBold;
		ws.getCell("Q" + (startCoal + 3)).alignment = textCenter;
		ws.getCell("Q" + (startCoal + 3)).border = borderThin;
		ws.getCell("Q" + (startCoal + 3)).fill = yellowHeader;

		ws.getCell("R" + (startCoal + 3)).value = "SOUTH";
		ws.getCell("R" + (startCoal + 3)).font = fontBold;
		ws.getCell("R" + (startCoal + 3)).alignment = textCenter;
		ws.getCell("R" + (startCoal + 3)).border = borderThin;
		ws.getCell("R" + (startCoal + 3)).fill = yellowHeader;

		ws.getCell("S" + (startCoal + 3)).value = "NORTH";
		ws.getCell("S" + (startCoal + 3)).font = fontBold;
		ws.getCell("S" + (startCoal + 3)).alignment = textCenter;
		ws.getCell("S" + (startCoal + 3)).border = borderThin;
		ws.getCell("S" + (startCoal + 3)).fill = yellowHeader;

		ws.getCell("T" + (startCoal + 3)).value = "SOUTH";
		ws.getCell("T" + (startCoal + 3)).font = fontBold;
		ws.getCell("T" + (startCoal + 3)).alignment = textCenter;
		ws.getCell("T" + (startCoal + 3)).border = borderThin;
		ws.getCell("T" + (startCoal + 3)).fill = yellowHeader;

		ws.getCell("U" + (startCoal + 3)).value = "NORTH";
		ws.getCell("U" + (startCoal + 3)).font = fontBold;
		ws.getCell("U" + (startCoal + 3)).alignment = textCenter;
		ws.getCell("U" + (startCoal + 3)).border = borderThin;
		ws.getCell("U" + (startCoal + 3)).fill = yellowHeader;

		let dataCoal1 = [
			{
				a: "1",
				b: "2",
				c: "3",
				d: "4",
				e: "5",
				f: "6",
				g: "7",
				h: "8",
				i: "9",
				j: "10",
				k: "11",
				l: "12",
				m: "13",
				n: "14",
				o: "15",
				p: "16",
				q: "17",
				r: "18",
				s: "19",
				t: "20",
			},
		];

		let startCoal1 = startCoal + 3;
		for (let i in dataCoal1) {
			startCoal1++;
			ws.getCell("B" + startCoal1).value = dataCoal1[i].a;
			ws.getCell("B" + startCoal1).alignment = textCenter;
			ws.getCell("B" + startCoal1).border = borderThin;

			ws.getCell("C" + startCoal1).value = dataCoal1[i].b;
			ws.getCell("C" + startCoal1).alignment = textCenter;
			ws.getCell("C" + startCoal1).border = borderThin;

			ws.getCell("D" + startCoal1).value = dataCoal1[i].c;
			ws.getCell("D" + startCoal1).alignment = textCenter;
			ws.getCell("D" + startCoal1).border = borderThin;

			ws.getCell("E" + startCoal1).value = dataCoal1[i].d;
			ws.getCell("E" + startCoal1).alignment = textCenter;
			ws.getCell("E" + startCoal1).border = borderThin;

			ws.getCell("F" + startCoal1).value = dataCoal1[i].e;
			ws.getCell("F" + startCoal1).alignment = textCenter;
			ws.getCell("F" + startCoal1).border = borderThin;

			ws.getCell("G" + startCoal1).value = dataCoal1[i].f;
			ws.getCell("G" + startCoal1).alignment = textCenter;
			ws.getCell("G" + startCoal1).border = borderThin;

			ws.getCell("H" + startCoal1).value = dataCoal1[i].g;
			ws.getCell("H" + startCoal1).alignment = textCenter;
			ws.getCell("H" + startCoal1).border = borderThin;

			ws.getCell("I" + startCoal1).value = dataCoal1[i].h;
			ws.getCell("I" + startCoal1).alignment = textCenter;
			ws.getCell("I" + startCoal1).border = borderThin;

			ws.getCell("J" + startCoal1).value = dataCoal1[i].i;
			ws.getCell("J" + startCoal1).alignment = textCenter;
			ws.getCell("J" + startCoal1).border = borderThin;

			ws.getCell("K" + startCoal1).value = dataCoal1[i].j;
			ws.getCell("K" + startCoal1).alignment = textCenter;
			ws.getCell("K" + startCoal1).border = borderThin;

			ws.getCell("L" + startCoal1).value = dataCoal1[i].k;
			ws.getCell("L" + startCoal1).alignment = textCenter;
			ws.getCell("L" + startCoal1).border = borderThin;

			ws.getCell("M" + startCoal1).value = dataCoal1[i].l;
			ws.getCell("M" + startCoal1).alignment = textCenter;
			ws.getCell("M" + startCoal1).border = borderThin;

			ws.getCell("N" + startCoal1).value = dataCoal1[i].m;
			ws.getCell("N" + startCoal1).alignment = textCenter;
			ws.getCell("N" + startCoal1).border = borderThin;

			ws.getCell("O" + startCoal1).value = dataCoal1[i].n;
			ws.getCell("O" + startCoal1).alignment = textCenter;
			ws.getCell("O" + startCoal1).border = borderThin;

			ws.getCell("P" + startCoal1).value = dataCoal1[i].o;
			ws.getCell("P" + startCoal1).alignment = textCenter;
			ws.getCell("P" + startCoal1).border = borderThin;

			ws.getCell("Q" + startCoal1).value = dataCoal1[i].p;
			ws.getCell("Q" + startCoal1).alignment = textCenter;
			ws.getCell("Q" + startCoal1).border = borderThin;

			ws.getCell("R" + startCoal1).value = dataCoal1[i].q;
			ws.getCell("R" + startCoal1).alignment = textCenter;
			ws.getCell("R" + startCoal1).border = borderThin;

			ws.getCell("S" + startCoal1).value = dataCoal1[i].r;
			ws.getCell("S" + startCoal1).alignment = textCenter;
			ws.getCell("S" + startCoal1).border = borderThin;

			ws.getCell("T" + startCoal1).value = dataCoal1[i].s;
			ws.getCell("T" + startCoal1).alignment = textCenter;
			ws.getCell("T" + startCoal1).border = borderThin;

			ws.getCell("U" + startCoal1).value = dataCoal1[i].t;
			ws.getCell("U" + startCoal1).alignment = textCenter;
			ws.getCell("U" + startCoal1).border = borderThin;
		}

		let startCoalHead2 = startCoal1 + 1;
		// coal stacker 2 header
		ws.mergeCells("B" + startCoalHead2 + ":" + "I" + startCoalHead2);
		ws.getCell("B" + startCoalHead2).value = "STACKER-2";
		ws.getCell("B" + startCoalHead2).font = fontBold;
		ws.getCell("B" + startCoalHead2).alignment = textCenter;
		ws.getCell("B" + startCoalHead2).border = borderThin;
		ws.getCell("B" + startCoalHead2).fill = cocoHeader;

		ws.mergeCells("J" + startCoalHead2 + ":" + "K" + (startCoalHead2 + 1));
		ws.getCell("J" + startCoalHead2).value = "TOTAL";
		ws.getCell("J" + startCoalHead2).font = fontBold;
		ws.getCell("J" + startCoalHead2).alignment = textCenter;
		ws.getCell("J" + startCoalHead2).border = borderThin;
		ws.getCell("J" + startCoalHead2).fill = cocoHeader;

		ws.mergeCells(
			"B" + (startCoalHead2 + 1) + ":" + "C" + (startCoalHead2 + 1)
		);
		ws.getCell("B" + (startCoalHead2 + 1)).value = "PRIMA";
		ws.getCell("B" + (startCoalHead2 + 1)).font = fontBold;
		ws.getCell("B" + (startCoalHead2 + 1)).alignment = textCenter;
		ws.getCell("B" + (startCoalHead2 + 1)).border = borderThin;
		ws.getCell("B" + (startCoalHead2 + 1)).fill = cocoHeader;
		//Head end olc 1 pinang header
		ws.mergeCells(
			"D" + (startCoalHead2 + 1) + ":" + "E" + (startCoalHead2 + 1)
		);
		ws.getCell("D" + (startCoalHead2 + 1)).value = "PINANG";
		ws.getCell("D" + (startCoalHead2 + 1)).font = fontBold;
		ws.getCell("D" + (startCoalHead2 + 1)).alignment = textCenter;
		ws.getCell("D" + (startCoalHead2 + 1)).border = borderThin;
		ws.getCell("D" + (startCoalHead2 + 1)).fill = cocoHeader;
		//Head end olc 1 melawan header
		ws.mergeCells(
			"F" + (startCoalHead2 + 1) + ":" + "G" + (startCoalHead2 + 1)
		);
		ws.getCell("F" + (startCoalHead2 + 1)).value = "MELAWAN";
		ws.getCell("F" + (startCoalHead2 + 1)).font = fontBold;
		ws.getCell("F" + (startCoalHead2 + 1)).alignment = textCenter;
		ws.getCell("F" + (startCoalHead2 + 1)).border = borderThin;
		ws.getCell("F" + (startCoalHead2 + 1)).fill = cocoHeader;
		//Head end olc 1 mlw42 header
		ws.mergeCells(
			"H" + (startCoalHead2 + 1) + ":" + "I" + (startCoalHead2 + 1)
		);
		ws.getCell("H" + (startCoalHead2 + 1)).value = "MLW42";
		ws.getCell("H" + (startCoalHead2 + 1)).font = fontBold;
		ws.getCell("H" + (startCoalHead2 + 1)).alignment = textCenter;
		ws.getCell("H" + (startCoalHead2 + 1)).border = borderThin;
		ws.getCell("H" + (startCoalHead2 + 1)).fill = cocoHeader;

		ws.getCell("B" + (startCoalHead2 + 2)).value = "SOUTH";
		ws.getCell("B" + (startCoalHead2 + 2)).font = fontBold;
		ws.getCell("B" + (startCoalHead2 + 2)).alignment = textCenter;
		ws.getCell("B" + (startCoalHead2 + 2)).border = borderThin;
		ws.getCell("B" + (startCoalHead2 + 2)).fill = cocoHeader;

		ws.getCell("C" + (startCoalHead2 + 2)).value = "NORTH";
		ws.getCell("C" + (startCoalHead2 + 2)).font = fontBold;
		ws.getCell("C" + (startCoalHead2 + 2)).alignment = textCenter;
		ws.getCell("C" + (startCoalHead2 + 2)).border = borderThin;
		ws.getCell("C" + (startCoalHead2 + 2)).fill = cocoHeader;

		ws.getCell("D" + (startCoalHead2 + 2)).value = "SOUTH";
		ws.getCell("D" + (startCoalHead2 + 2)).font = fontBold;
		ws.getCell("D" + (startCoalHead2 + 2)).alignment = textCenter;
		ws.getCell("D" + (startCoalHead2 + 2)).border = borderThin;
		ws.getCell("D" + (startCoalHead2 + 2)).fill = cocoHeader;

		ws.getCell("E" + (startCoalHead2 + 2)).value = "NORTH";
		ws.getCell("E" + (startCoalHead2 + 2)).font = fontBold;
		ws.getCell("E" + (startCoalHead2 + 2)).alignment = textCenter;
		ws.getCell("E" + (startCoalHead2 + 2)).border = borderThin;
		ws.getCell("E" + (startCoalHead2 + 2)).fill = cocoHeader;

		ws.getCell("F" + (startCoalHead2 + 2)).value = "SOUTH";
		ws.getCell("F" + (startCoalHead2 + 2)).font = fontBold;
		ws.getCell("F" + (startCoalHead2 + 2)).alignment = textCenter;
		ws.getCell("F" + (startCoalHead2 + 2)).border = borderThin;
		ws.getCell("F" + (startCoalHead2 + 2)).fill = cocoHeader;

		ws.getCell("G" + (startCoalHead2 + 2)).value = "NORTH";
		ws.getCell("G" + (startCoalHead2 + 2)).font = fontBold;
		ws.getCell("G" + (startCoalHead2 + 2)).alignment = textCenter;
		ws.getCell("G" + (startCoalHead2 + 2)).border = borderThin;
		ws.getCell("G" + (startCoalHead2 + 2)).fill = cocoHeader;

		ws.getCell("H" + (startCoalHead2 + 2)).value = "SOUTH";
		ws.getCell("H" + (startCoalHead2 + 2)).font = fontBold;
		ws.getCell("H" + (startCoalHead2 + 2)).alignment = textCenter;
		ws.getCell("H" + (startCoalHead2 + 2)).border = borderThin;
		ws.getCell("H" + (startCoalHead2 + 2)).fill = cocoHeader;

		ws.getCell("I" + (startCoalHead2 + 2)).value = "NORTH";
		ws.getCell("I" + (startCoalHead2 + 2)).font = fontBold;
		ws.getCell("I" + (startCoalHead2 + 2)).alignment = textCenter;
		ws.getCell("I" + (startCoalHead2 + 2)).border = borderThin;
		ws.getCell("I" + (startCoalHead2 + 2)).fill = cocoHeader;

		ws.getCell("J" + (startCoalHead2 + 2)).value = "SOUTH";
		ws.getCell("J" + (startCoalHead2 + 2)).font = fontBold;
		ws.getCell("J" + (startCoalHead2 + 2)).alignment = textCenter;
		ws.getCell("J" + (startCoalHead2 + 2)).border = borderThin;
		ws.getCell("J" + (startCoalHead2 + 2)).fill = cocoHeader;

		ws.getCell("K" + (startCoalHead2 + 2)).value = "NORTH";
		ws.getCell("K" + (startCoalHead2 + 2)).font = fontBold;
		ws.getCell("K" + (startCoalHead2 + 2)).alignment = textCenter;
		ws.getCell("K" + (startCoalHead2 + 2)).border = borderThin;
		ws.getCell("K" + (startCoalHead2 + 2)).fill = cocoHeader;

		// COAL RECLAIMING 2 HEADER
		ws.mergeCells("L" + startCoalHead2 + ":" + "S" + startCoalHead2);
		ws.getCell("L" + startCoalHead2).value = "RECLAIMER-2";
		ws.getCell("L" + startCoalHead2).font = fontBold;
		ws.getCell("L" + startCoalHead2).alignment = textCenter;
		ws.getCell("L" + startCoalHead2).border = borderThin;
		ws.getCell("L" + startCoalHead2).fill = yellowHeader;

		ws.mergeCells("T" + startCoalHead2 + ":" + "U" + (startCoalHead2 + 1));
		ws.getCell("T" + startCoalHead2).value = "TOTAL";
		ws.getCell("T" + startCoalHead2).font = fontBold;
		ws.getCell("T" + startCoalHead2).alignment = textCenter;
		ws.getCell("T" + startCoalHead2).border = borderThin;
		ws.getCell("T" + startCoalHead2).fill = yellowHeader;

		ws.mergeCells(
			"L" + (startCoalHead2 + 1) + ":" + "M" + (startCoalHead2 + 1)
		);
		ws.getCell("L" + (startCoalHead2 + 1)).value = "PRIMA";
		ws.getCell("L" + (startCoalHead2 + 1)).font = fontBold;
		ws.getCell("L" + (startCoalHead2 + 1)).alignment = textCenter;
		ws.getCell("L" + (startCoalHead2 + 1)).border = borderThin;
		ws.getCell("L" + (startCoalHead2 + 1)).fill = yellowHeader;
		//Head end olc 1 pinang header
		ws.mergeCells(
			"N" + (startCoalHead2 + 1) + ":" + "O" + (startCoalHead2 + 1)
		);
		ws.getCell("N" + (startCoalHead2 + 1)).value = "PINANG";
		ws.getCell("N" + (startCoalHead2 + 1)).font = fontBold;
		ws.getCell("N" + (startCoalHead2 + 1)).alignment = textCenter;
		ws.getCell("N" + (startCoalHead2 + 1)).border = borderThin;
		ws.getCell("N" + (startCoalHead2 + 1)).fill = yellowHeader;
		//Head end olc 1 melawan header
		ws.mergeCells(
			"P" + (startCoalHead2 + 1) + ":" + "Q" + (startCoalHead2 + 1)
		);
		ws.getCell("P" + (startCoalHead2 + 1)).value = "MELAWAN";
		ws.getCell("P" + (startCoalHead2 + 1)).font = fontBold;
		ws.getCell("P" + (startCoalHead2 + 1)).alignment = textCenter;
		ws.getCell("P" + (startCoalHead2 + 1)).border = borderThin;
		ws.getCell("P" + (startCoalHead2 + 1)).fill = yellowHeader;
		//Head end olc 1 mlw42 header
		ws.mergeCells(
			"R" + (startCoalHead2 + 1) + ":" + "S" + (startCoalHead2 + 1)
		);
		ws.getCell("R" + (startCoalHead2 + 1)).value = "MLW42";
		ws.getCell("R" + (startCoalHead2 + 1)).font = fontBold;
		ws.getCell("R" + (startCoalHead2 + 1)).alignment = textCenter;
		ws.getCell("R" + (startCoalHead2 + 1)).border = borderThin;
		ws.getCell("R" + (startCoalHead2 + 1)).fill = yellowHeader;

		ws.getCell("L" + (startCoalHead2 + 2)).value = "SOUTH";
		ws.getCell("L" + (startCoalHead2 + 2)).font = fontBold;
		ws.getCell("L" + (startCoalHead2 + 2)).alignment = textCenter;
		ws.getCell("L" + (startCoalHead2 + 2)).border = borderThin;
		ws.getCell("L" + (startCoalHead2 + 2)).fill = yellowHeader;

		ws.getCell("M" + (startCoalHead2 + 2)).value = "NORTH";
		ws.getCell("M" + (startCoalHead2 + 2)).font = fontBold;
		ws.getCell("M" + (startCoalHead2 + 2)).alignment = textCenter;
		ws.getCell("M" + (startCoalHead2 + 2)).border = borderThin;
		ws.getCell("M" + (startCoalHead2 + 2)).fill = yellowHeader;

		ws.getCell("N" + (startCoalHead2 + 2)).value = "SOUTH";
		ws.getCell("N" + (startCoalHead2 + 2)).font = fontBold;
		ws.getCell("N" + (startCoalHead2 + 2)).alignment = textCenter;
		ws.getCell("N" + (startCoalHead2 + 2)).border = borderThin;
		ws.getCell("N" + (startCoalHead2 + 2)).fill = yellowHeader;

		ws.getCell("O" + (startCoalHead2 + 2)).value = "NORTH";
		ws.getCell("O" + (startCoalHead2 + 2)).font = fontBold;
		ws.getCell("O" + (startCoalHead2 + 2)).alignment = textCenter;
		ws.getCell("O" + (startCoalHead2 + 2)).border = borderThin;
		ws.getCell("O" + (startCoalHead2 + 2)).fill = yellowHeader;

		ws.getCell("P" + (startCoalHead2 + 2)).value = "SOUTH";
		ws.getCell("P" + (startCoalHead2 + 2)).font = fontBold;
		ws.getCell("P" + (startCoalHead2 + 2)).alignment = textCenter;
		ws.getCell("P" + (startCoalHead2 + 2)).border = borderThin;
		ws.getCell("P" + (startCoalHead2 + 2)).fill = yellowHeader;

		ws.getCell("Q" + (startCoalHead2 + 2)).value = "NORTH";
		ws.getCell("Q" + (startCoalHead2 + 2)).font = fontBold;
		ws.getCell("Q" + (startCoalHead2 + 2)).alignment = textCenter;
		ws.getCell("Q" + (startCoalHead2 + 2)).border = borderThin;
		ws.getCell("Q" + (startCoalHead2 + 2)).fill = yellowHeader;

		ws.getCell("R" + (startCoalHead2 + 2)).value = "SOUTH";
		ws.getCell("R" + (startCoalHead2 + 2)).font = fontBold;
		ws.getCell("R" + (startCoalHead2 + 2)).alignment = textCenter;
		ws.getCell("R" + (startCoalHead2 + 2)).border = borderThin;
		ws.getCell("R" + (startCoalHead2 + 2)).fill = yellowHeader;

		ws.getCell("S" + (startCoalHead2 + 2)).value = "NORTH";
		ws.getCell("S" + (startCoalHead2 + 2)).font = fontBold;
		ws.getCell("S" + (startCoalHead2 + 2)).alignment = textCenter;
		ws.getCell("S" + (startCoalHead2 + 2)).border = borderThin;
		ws.getCell("S" + (startCoalHead2 + 2)).fill = yellowHeader;

		ws.getCell("T" + (startCoalHead2 + 2)).value = "SOUTH";
		ws.getCell("T" + (startCoalHead2 + 2)).font = fontBold;
		ws.getCell("T" + (startCoalHead2 + 2)).alignment = textCenter;
		ws.getCell("T" + (startCoalHead2 + 2)).border = borderThin;
		ws.getCell("T" + (startCoalHead2 + 2)).fill = yellowHeader;

		ws.getCell("U" + (startCoalHead2 + 2)).value = "NORTH";
		ws.getCell("U" + (startCoalHead2 + 2)).font = fontBold;
		ws.getCell("U" + (startCoalHead2 + 2)).alignment = textCenter;
		ws.getCell("U" + (startCoalHead2 + 2)).border = borderThin;
		ws.getCell("U" + (startCoalHead2 + 2)).fill = yellowHeader;

		let dataCoal2 = [
			{
				a: "1",
				b: "2",
				c: "3",
				d: "4",
				e: "5",
				f: "6",
				g: "7",
				h: "8",
				i: "9",
				j: "10",
				k: "11",
				l: "12",
				m: "13",
				n: "14",
				o: "15",
				p: "16",
				q: "17",
				r: "18",
				s: "19",
				t: "20",
			},
		];

		let startCoal2 = startCoalHead2 + 2;
		for (let i in dataCoal2) {
			startCoal2++;
			ws.getCell("B" + startCoal2).value = dataCoal2[i].a;
			ws.getCell("B" + startCoal2).alignment = textCenter;
			ws.getCell("B" + startCoal2).border = borderThin;

			ws.getCell("C" + startCoal2).value = dataCoal2[i].b;
			ws.getCell("C" + startCoal2).alignment = textCenter;
			ws.getCell("C" + startCoal2).border = borderThin;

			ws.getCell("D" + startCoal2).value = dataCoal2[i].c;
			ws.getCell("D" + startCoal2).alignment = textCenter;
			ws.getCell("D" + startCoal2).border = borderThin;

			ws.getCell("E" + startCoal2).value = dataCoal2[i].d;
			ws.getCell("E" + startCoal2).alignment = textCenter;
			ws.getCell("E" + startCoal2).border = borderThin;

			ws.getCell("F" + startCoal2).value = dataCoal2[i].e;
			ws.getCell("F" + startCoal2).alignment = textCenter;
			ws.getCell("F" + startCoal2).border = borderThin;

			ws.getCell("G" + startCoal2).value = dataCoal2[i].f;
			ws.getCell("G" + startCoal2).alignment = textCenter;
			ws.getCell("G" + startCoal2).border = borderThin;

			ws.getCell("H" + startCoal2).value = dataCoal2[i].g;
			ws.getCell("H" + startCoal2).alignment = textCenter;
			ws.getCell("H" + startCoal2).border = borderThin;

			ws.getCell("I" + startCoal2).value = dataCoal2[i].h;
			ws.getCell("I" + startCoal2).alignment = textCenter;
			ws.getCell("I" + startCoal2).border = borderThin;

			ws.getCell("J" + startCoal2).value = dataCoal2[i].i;
			ws.getCell("J" + startCoal2).alignment = textCenter;
			ws.getCell("J" + startCoal2).border = borderThin;

			ws.getCell("K" + startCoal2).value = dataCoal2[i].j;
			ws.getCell("K" + startCoal2).alignment = textCenter;
			ws.getCell("K" + startCoal2).border = borderThin;

			ws.getCell("L" + startCoal2).value = dataCoal2[i].k;
			ws.getCell("L" + startCoal2).alignment = textCenter;
			ws.getCell("L" + startCoal2).border = borderThin;

			ws.getCell("M" + startCoal2).value = dataCoal2[i].l;
			ws.getCell("M" + startCoal2).alignment = textCenter;
			ws.getCell("M" + startCoal2).border = borderThin;

			ws.getCell("N" + startCoal2).value = dataCoal2[i].m;
			ws.getCell("N" + startCoal2).alignment = textCenter;
			ws.getCell("N" + startCoal2).border = borderThin;

			ws.getCell("O" + startCoal2).value = dataCoal2[i].n;
			ws.getCell("O" + startCoal2).alignment = textCenter;
			ws.getCell("O" + startCoal2).border = borderThin;

			ws.getCell("P" + startCoal2).value = dataCoal2[i].o;
			ws.getCell("P" + startCoal2).alignment = textCenter;
			ws.getCell("P" + startCoal2).border = borderThin;

			ws.getCell("Q" + startCoal2).value = dataCoal2[i].p;
			ws.getCell("Q" + startCoal2).alignment = textCenter;
			ws.getCell("Q" + startCoal2).border = borderThin;

			ws.getCell("R" + startCoal2).value = dataCoal2[i].q;
			ws.getCell("R" + startCoal2).alignment = textCenter;
			ws.getCell("R" + startCoal2).border = borderThin;

			ws.getCell("S" + startCoal2).value = dataCoal2[i].r;
			ws.getCell("S" + startCoal2).alignment = textCenter;
			ws.getCell("S" + startCoal2).border = borderThin;

			ws.getCell("T" + startCoal2).value = dataCoal2[i].s;
			ws.getCell("T" + startCoal2).alignment = textCenter;
			ws.getCell("T" + startCoal2).border = borderThin;

			ws.getCell("U" + startCoal2).value = dataCoal2[i].t;
			ws.getCell("U" + startCoal2).alignment = textCenter;
			ws.getCell("U" + startCoal2).border = borderThin;
		}

		let startTotalStacking = startCoal2 + 1;
		ws.mergeCells("B" + startTotalStacking + ":" + "I" + startTotalStacking);
		ws.getCell("B" + startTotalStacking).value = "TOTAL STACKING";
		ws.getCell("B" + startTotalStacking).font = fontBold;
		ws.getCell("B" + startTotalStacking).border = borderThin;
		ws.getCell("B" + startTotalStacking).alignment = textRight;
		// value total stacking
		ws.mergeCells("J" + startTotalStacking + ":" + "K" + startTotalStacking);
		ws.getCell("J" + startTotalStacking).value = "1200";
		ws.getCell("J" + startTotalStacking).font = fontBold;
		ws.getCell("J" + startTotalStacking).border = borderThin;
		ws.getCell("J" + startTotalStacking).alignment = textRight;
		//OLC 1 BYPASS
		ws.mergeCells("L" + startTotalStacking + ":" + "S" + startTotalStacking);
		ws.getCell("L" + startTotalStacking).value = "OLC#1 BYPASS";
		ws.getCell("L" + startTotalStacking).font = fontBold;
		ws.getCell("L" + startTotalStacking).border = borderThin;
		ws.getCell("L" + startTotalStacking).alignment = textCenter;
		ws.getCell("L" + startTotalStacking).fill = yellowHeader;

		ws.mergeCells(
			"T" + startTotalStacking + ":" + "U" + (startTotalStacking + 1)
		);
		ws.getCell("T" + startTotalStacking).value = "TOTAL";
		ws.getCell("T" + startTotalStacking).font = fontBold;
		ws.getCell("T" + startTotalStacking).border = borderThin;
		ws.getCell("T" + startTotalStacking).alignment = textCenter;
		ws.getCell("T" + startTotalStacking).fill = yellowHeader;

		ws.mergeCells(
			"L" + (startTotalStacking + 1) + ":" + "M" + (startTotalStacking + 1)
		);
		ws.getCell("L" + (startTotalStacking + 1)).value = "PRIMA";
		ws.getCell("L" + (startTotalStacking + 1)).font = fontBold;
		ws.getCell("L" + (startTotalStacking + 1)).alignment = textCenter;
		ws.getCell("L" + (startTotalStacking + 1)).border = borderThin;
		ws.getCell("L" + (startTotalStacking + 1)).fill = yellowHeader;
		//Head end olc 1 pinang header
		ws.mergeCells(
			"N" + (startTotalStacking + 1) + ":" + "O" + (startTotalStacking + 1)
		);
		ws.getCell("N" + (startTotalStacking + 1)).value = "PINANG";
		ws.getCell("N" + (startTotalStacking + 1)).font = fontBold;
		ws.getCell("N" + (startTotalStacking + 1)).alignment = textCenter;
		ws.getCell("N" + (startTotalStacking + 1)).border = borderThin;
		ws.getCell("N" + (startTotalStacking + 1)).fill = yellowHeader;
		//Head end olc 1 melawan header
		ws.mergeCells(
			"P" + (startTotalStacking + 1) + ":" + "Q" + (startTotalStacking + 1)
		);
		ws.getCell("P" + (startTotalStacking + 1)).value = "MELAWAN";
		ws.getCell("P" + (startTotalStacking + 1)).font = fontBold;
		ws.getCell("P" + (startTotalStacking + 1)).alignment = textCenter;
		ws.getCell("P" + (startTotalStacking + 1)).border = borderThin;
		ws.getCell("P" + (startTotalStacking + 1)).fill = yellowHeader;
		//Head end olc 1 mlw42 header
		ws.mergeCells(
			"R" + (startTotalStacking + 1) + ":" + "S" + (startTotalStacking + 1)
		);
		ws.getCell("R" + (startTotalStacking + 1)).value = "MLW42";
		ws.getCell("R" + (startTotalStacking + 1)).font = fontBold;
		ws.getCell("R" + (startTotalStacking + 1)).alignment = textCenter;
		ws.getCell("R" + (startTotalStacking + 1)).border = borderThin;
		ws.getCell("R" + (startTotalStacking + 1)).fill = yellowHeader;

		let dataOLC1 = [
			{
				a: "100",
				b: "200",
				c: "300",
				d: "400",
				e: "1000",
			},
		];

		let startDataOLC1 = startTotalStacking + 1;

		for (let i in dataOLC1) {
			startDataOLC1++;
			ws.mergeCells("L" + startDataOLC1 + ":" + "M" + startDataOLC1);
			ws.getCell("L" + startDataOLC1).value = dataOLC1[i].a;
			ws.getCell("L" + startDataOLC1).alignment = textCenter;
			ws.getCell("L" + startDataOLC1).border = borderThin;
			//Head end olc 1 pinang header
			ws.mergeCells("N" + startDataOLC1 + ":" + "O" + startDataOLC1);
			ws.getCell("N" + startDataOLC1).value = dataOLC1[i].b;
			ws.getCell("N" + startDataOLC1).alignment = textCenter;
			ws.getCell("N" + startDataOLC1).border = borderThin;
			//Head end olc 1 melawan header
			ws.mergeCells("P" + startDataOLC1 + ":" + "Q" + startDataOLC1);
			ws.getCell("P" + startDataOLC1).value = dataOLC1[i].c;
			ws.getCell("P" + startDataOLC1).alignment = textCenter;
			ws.getCell("P" + startDataOLC1).border = borderThin;
			//Head end olc 1 mlw42 header
			ws.mergeCells("R" + startDataOLC1 + ":" + "S" + startDataOLC1);
			ws.getCell("R" + startDataOLC1).value = dataOLC1[i].d;
			ws.getCell("R" + startDataOLC1).alignment = textCenter;
			ws.getCell("R" + startDataOLC1).border = borderThin;
			//tptal
			ws.mergeCells("T" + startDataOLC1 + ":" + "U" + startDataOLC1);
			ws.getCell("T" + startDataOLC1).value = dataOLC1[i].e;
			ws.getCell("T" + startDataOLC1).alignment = textCenter;
			ws.getCell("T" + startDataOLC1).border = borderThin;
		}

		let startOLC2 = startDataOLC1 + 1;

		//OLC 1 BYPASS
		ws.mergeCells("L" + startOLC2 + ":" + "S" + startOLC2);
		ws.getCell("L" + startOLC2).value = "OLC#2 BYPASS";
		ws.getCell("L" + startOLC2).font = fontBold;
		ws.getCell("L" + startOLC2).border = borderThin;
		ws.getCell("L" + startOLC2).alignment = textCenter;
		ws.getCell("L" + startOLC2).fill = yellowHeader;

		ws.mergeCells("T" + startOLC2 + ":" + "U" + (startOLC2 + 1));
		ws.getCell("T" + startOLC2).value = "TOTAL";
		ws.getCell("T" + startOLC2).font = fontBold;
		ws.getCell("T" + startOLC2).border = borderThin;
		ws.getCell("T" + startOLC2).alignment = textCenter;
		ws.getCell("T" + startOLC2).fill = yellowHeader;

		ws.mergeCells("L" + (startOLC2 + 1) + ":" + "M" + (startOLC2 + 1));
		ws.getCell("L" + (startOLC2 + 1)).value = "PRIMA";
		ws.getCell("L" + (startOLC2 + 1)).font = fontBold;
		ws.getCell("L" + (startOLC2 + 1)).alignment = textCenter;
		ws.getCell("L" + (startOLC2 + 1)).border = borderThin;
		ws.getCell("L" + (startOLC2 + 1)).fill = yellowHeader;
		//Head end olc 1 pinang header
		ws.mergeCells("N" + (startOLC2 + 1) + ":" + "O" + (startOLC2 + 1));
		ws.getCell("N" + (startOLC2 + 1)).value = "PINANG";
		ws.getCell("N" + (startOLC2 + 1)).font = fontBold;
		ws.getCell("N" + (startOLC2 + 1)).alignment = textCenter;
		ws.getCell("N" + (startOLC2 + 1)).border = borderThin;
		ws.getCell("N" + (startOLC2 + 1)).fill = yellowHeader;
		//Head end olc 1 melawan header
		ws.mergeCells("P" + (startOLC2 + 1) + ":" + "Q" + (startOLC2 + 1));
		ws.getCell("P" + (startOLC2 + 1)).value = "MELAWAN";
		ws.getCell("P" + (startOLC2 + 1)).font = fontBold;
		ws.getCell("P" + (startOLC2 + 1)).alignment = textCenter;
		ws.getCell("P" + (startOLC2 + 1)).border = borderThin;
		ws.getCell("P" + (startOLC2 + 1)).fill = yellowHeader;
		//Head end olc 1 mlw42 header
		ws.mergeCells("R" + (startOLC2 + 1) + ":" + "S" + (startOLC2 + 1));
		ws.getCell("R" + (startOLC2 + 1)).value = "MLW42";
		ws.getCell("R" + (startOLC2 + 1)).font = fontBold;
		ws.getCell("R" + (startOLC2 + 1)).alignment = textCenter;
		ws.getCell("R" + (startOLC2 + 1)).border = borderThin;
		ws.getCell("R" + (startOLC2 + 1)).fill = yellowHeader;

		let dataOLC2 = [
			{
				a: "1000",
				b: "2000",
				c: "3000",
				d: "4000",
				e: "10000",
			},
		];

		let startDataOLC2 = startOLC2 + 1;

		for (let i in dataOLC2) {
			startDataOLC2++;
			ws.mergeCells("L" + startDataOLC2 + ":" + "M" + startDataOLC2);
			ws.getCell("L" + startDataOLC2).value = dataOLC2[i].a;
			ws.getCell("L" + startDataOLC2).alignment = textCenter;
			ws.getCell("L" + startDataOLC2).border = borderThin;
			//Head end olc 1 pinang header
			ws.mergeCells("N" + startDataOLC2 + ":" + "O" + startDataOLC2);
			ws.getCell("N" + startDataOLC2).value = dataOLC2[i].b;
			ws.getCell("N" + startDataOLC2).alignment = textCenter;
			ws.getCell("N" + startDataOLC2).border = borderThin;
			//Head end olc 1 melawan header
			ws.mergeCells("P" + startDataOLC2 + ":" + "Q" + startDataOLC2);
			ws.getCell("P" + startDataOLC2).value = dataOLC2[i].c;
			ws.getCell("P" + startDataOLC2).alignment = textCenter;
			ws.getCell("P" + startDataOLC2).border = borderThin;
			//Head end olc 1 mlw42 header
			ws.mergeCells("R" + startDataOLC2 + ":" + "S" + startDataOLC2);
			ws.getCell("R" + startDataOLC2).value = dataOLC2[i].d;
			ws.getCell("R" + startDataOLC2).alignment = textCenter;
			ws.getCell("R" + startDataOLC2).border = borderThin;
			//tptal
			ws.mergeCells("T" + startDataOLC2 + ":" + "U" + startDataOLC2);
			ws.getCell("T" + startDataOLC2).value = dataOLC2[i].e;
			ws.getCell("T" + startDataOLC2).alignment = textCenter;
			ws.getCell("T" + startDataOLC2).border = borderThin;
		}

		let startTotalReclaiming = startDataOLC2 + 1;

		ws.mergeCells(
			"L" + startTotalReclaiming + ":" + "S" + startTotalReclaiming
		);
		ws.getCell("L" + startTotalReclaiming).value = "TOTAL RECLAIMING";
		ws.getCell("L" + startTotalReclaiming).alignment = textCenter;
		ws.getCell("L" + startTotalReclaiming).border = borderThin;
		ws.getCell("L" + startTotalReclaiming).font = fontBold;
		//data total reclaiming
		ws.mergeCells(
			"T" + startTotalReclaiming + ":" + "U" + startTotalReclaiming
		);
		ws.getCell("T" + startTotalReclaiming).value = "120000";
		ws.getCell("T" + startTotalReclaiming).alignment = textCenter;
		ws.getCell("T" + startTotalReclaiming).border = borderThin;
		ws.getCell("T" + startTotalReclaiming).font = fontBold;

		let datePrint = `${moment().format("YYYYMMDD")}_${moment().format(
			"HHmmss"
		)}`;
		const direct = "/tmp";
		await wb.xlsx
			.writeFile(`${direct}/WEIGHER_REPORT_${datePrint}.xlsx`)
			.then(() => {
				res.download(
					`${direct}/WEIGHER_REPORT_${datePrint}.xlsx`,
					`WEIGHER_REPORT_${datePrint}.xlsx`,
					(err) => {
						if (err) {
							console.log(err);
						} else {
							fs.unlinkSync(`${direct}/WEIGHER_REPORT_${datePrint}.xlsx`);
						}
					}
				);
			});
	} catch (e) {
		console.log(e);
		res.json({ status: "failed", reason: e });
	}
};

module.exports = {
	exportExcel,
};
