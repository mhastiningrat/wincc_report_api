const grayHeader = {
	type: "pattern",
	pattern: "solid",
	fgColor: { argb: "d4d9d5" },
};

const yellowHeader = {
	type: "pattern",
	pattern: "solid",
	fgColor: { argb: "f7e91e" },
};

const cocoHeader = {
	type: "pattern",
	pattern: "solid",
	fgColor: { argb: "fac289" },
};

const leafHeader = {
	type: "pattern",
	pattern: "solid",
	fgColor: { argb: "2bb550" },
};

const oceanHeader = {
	type: "pattern",
	pattern: "solid",
	fgColor: { argb: "21dffc" },
};

const oceanMudaHeader = {
	type: "pattern",
	pattern: "solid",
	fgColor: { argb: "b7e6ed" },
};

const toscaHeader = {
	type: "pattern",
	pattern: "solid",
	fgColor: { argb: "c6f5e6" },
};

const fontBold = {
	bold: true,
};

const fontBolder = {
	bold: true,
	size: 22,
};

const textCenter = { vertical: "middle", horizontal: "center" };
const textLeft = { vertical: "middle", horizontal: "left" };
const textRight = { vertical: "middle", horizontal: "right" };

const borderBold = {
	top: { style: "medium", color: { argb: "000000" } },
	left: { style: "medium", color: { argb: "000000" } },
	bottom: { style: "medium", color: { argb: "000000" } },
	right: { style: "medium", color: { argb: "000000" } },
};

const borderThin = {
	top: { style: "thin", color: { argb: "000000" } },
	left: { style: "thin", color: { argb: "000000" } },
	bottom: { style: "thin", color: { argb: "000000" } },
	right: { style: "thin", color: { argb: "000000" } },
};

module.exports = {
	grayHeader,
	yellowHeader,
	leafHeader,
	oceanHeader,
	fontBold,
	textCenter,
	textLeft,
	textRight,
	borderBold,
	borderThin,
	cocoHeader,
	fontBolder,
	oceanMudaHeader,
	toscaHeader,
};
