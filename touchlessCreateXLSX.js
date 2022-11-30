// Import XLSX package
const XLSX = require('xlsx-js-style');
const data = "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";
let workbook = XLSX.read(data);
const sheetName = workbook.SheetNames[0];

const colors = {
    darkBlue: "285FE5",
    skyBlue: "CCECFF",
    babyBlue: "B4C6E7",
    lightBlue: "D9E1F2",
    cactus: "A9D08E",
    peach: "F4B084",
    lightTan: "FCE4D6",
    tan: "F8CBAD",
    red: "EA3323",
    gray: "C9C9C9",
    silver: "DBDBDB",
    black: "000000",
    white: "FFFFFF",
    rain: "D9D9D9",
    cloudy: "BFBFBF",
    pink: "F4C2C2",
};

const createCell = (cell) => {
    workbook.Sheets[sheetName][cell] = { t: "s", v: "", s: {} };
};

const checkCell = (cell) => {
    // If the cell doesn't exist yet, create it
    if (!workbook.Sheets[sheetName][cell]) createCell(cell);
    // If the cell doesn't have a style yet, create it
    if (!workbook.Sheets[sheetName][cell].s) workbook.Sheets[sheetName][cell].s = {};
};

const addBorder = (cell, direction, weight="medium") => {
    checkCell(cell);
    workbook.Sheets[sheetName][cell].s.border = {
        ...workbook.Sheets[sheetName][cell].s?.border,
        [direction]: {
            style: weight,
            color: { rgb: colors.black },
        }
    };
};

const addBgColor = (cell, color) => {
    checkCell(cell);
    workbook.Sheets[sheetName][cell].s.fill = {
        fgColor: { rgb: color }
    };
};

const addAlignment = (cell, alignment, value) => {
    checkCell(cell);
    workbook.Sheets[sheetName][cell].s.alignment = {
        ...workbook.Sheets[sheetName][cell].s?.alignment,
        [alignment]: value
    };
};

const addFont = (cell, fontProp, value) => {
    checkCell(cell);
    workbook.Sheets[sheetName][cell].s.font = {
        ...workbook.Sheets[sheetName][cell].s?.font,
        [fontProp]: value
    };
};

// Make background color of cell A1 hex #285FE5 and text color white
addBgColor("A1", colors.darkBlue);
addFont("A1", "color", { rgb: colors.white });

// Make cell A3 bold and text color #285FE5 and font size 28
addFont("A3", "bold", true);
addFont("A3", "color", { rgb: colors.darkBlue });
addFont("A3", "sz", 28);

// Call addBorder function to add top borders to cells A5 - Q5
cellFive = ['A5', 'B5', 'C5', 'D5', 'E5', 'F5', 'G5', 'H5', 'I5', 'J5', 'K5', 'L5', 'M5', 'N5', 'O5', 'P5', 'Q5'];
cellFive.forEach(cell => { 
    addBorder(cell, 'bottom');
    addBorder(cell, 'top');
});

// Make background color of cell C5 and L5 cactus
addBgColor('C5', colors.cactus);
addBgColor('L5', colors.cactus);

// Make cell A6 bold
addFont('A6', 'bold', true);

// Make cell C6 right aligned
addAlignment('C6', 'horizontal', 'right');

// Make cell F6 background color peach
addBgColor('F6', colors.peach);

// Make cell I6 right aligned
addAlignment('I6', 'horizontal', 'right');

// Make cell L6 background color peach
addBgColor('L6', colors.peach);

// Make border top and bottom for cells C6 - N6
const cellSix = ['C6', 'D6', 'E6', 'F6', 'G6', 'H6', 'I6', 'J6', 'K6', 'L6', 'M6', 'N6'];
cellSix.forEach((cell, idx) => { 
    if (idx === 0) {
        addBorder(cell, 'left');
    }
    if (idx === cellSix.length - 1) {
        addBorder(cell, 'right');
    }
    addBorder(cell, 'bottom'); 
    addBorder(cell, 'top');
});

// Make cell C7, K7, P7 background color cactus
addBgColor('C7', colors.cactus);
addBgColor('K7', colors.cactus);
addBgColor('P7', colors.cactus);

// Make light border bottom for cells A7 - Q7
const cellSeven = ['A7', 'B7', 'C7', 'D7', 'E7', 'F7', 'G7', 'H7', 'I7', 'J7', 'K7', 'L7', 'M7', 'N7', 'O7', 'P7', 'Q7'];
cellSeven.forEach(cell => {
    addBorder(cell, 'bottom', 'thin');
});

// Make cell C8, K8 background color cactus and P8 background color red
addBgColor('C8', colors.cactus);
addBgColor('K8', colors.cactus);
addBgColor('P8', colors.red);

// Make light border bottom for cells A8 - Q8
const cellEight = ['A8', 'B8', 'C8', 'D8', 'E8', 'F8', 'G8', 'H8', 'I8', 'J8', 'K8', 'L8', 'M8', 'N8', 'O8', 'P8', 'Q8'];
cellEight.forEach(cell => {
    addBorder(cell, 'bottom', 'thin');
});

// Make cell C9, K9 background color cactus and P9 background color red
addBgColor('C9', colors.cactus);
addBgColor('K9', colors.cactus);
addBgColor('P9', colors.red);

// Add medium border top for cells A10 - Q10
const cellTen = ['A10', 'B10', 'C10', 'D10', 'E10', 'F10', 'G10', 'H10', 'I10', 'J10', 'K10', 'L10', 'M10', 'N10', 'O10', 'P10', 'Q10'];
cellTen.forEach((cell, idx) => {
    if (idx === 0) {
        addFont(cell, 'bold', true);
        addBgColor(cell, colors.silver);
    } else if (idx % 2 !== 0) {
        // If idx is odd, give border left, center text size 8 bold
        addBorder(cell, 'left', 'thin');
        addBorder(cell, 'bottom', 'thin');
        addAlignment(cell, 'horizontal', 'center');
        addFont(cell, 'sz', 8);
        addFont(cell, 'bold', true);
        addBgColor(cell, colors.gray);
    } else {
        // If idx is even number, excluding 0, add light border right and bottom
        addBorder(cell, 'right', 'thin');
        addBorder(cell, 'bottom', 'thin');

    }
    
    addBorder(cell, 'top');
});

// Cells A11 - Q11
const cellEleven = ['A11', 'B11', 'C11', 'D11', 'E11', 'F11', 'G11', 'H11', 'I11', 'J11', 'K11', 'L11', 'M11', 'N11', 'O11', 'P11', 'Q11'];
cellEleven.forEach((cell, idx) => {
    if (idx > 0) {
        // Border left right bottom
        addBorder(cell, 'left', 'thin');
        addBorder(cell, 'right', 'thin');
        addBorder(cell, 'bottom', 'thin');
        // Center text size 8 bold
        addAlignment(cell, 'horizontal', 'center');
        addFont(cell, 'sz', 8);
        addFont(cell, 'bold', true);
        // Background color gray
        addBgColor(cell, colors.gray);
    } else {
        addBgColor(cell, colors.silver);
    }
});

// Cells A12 - Q12
const cellTwelve = ['A12', 'B12', 'C12', 'D12', 'E12', 'F12', 'G12', 'H12', 'I12', 'J12', 'K12', 'L12', 'M12', 'N12', 'O12', 'P12', 'Q12'];
cellTwelve.forEach((cell, idx) => {
    if (idx > 0) {
        // Border left right
        addBorder(cell, 'left', 'thin');
        addBorder(cell, 'right', 'thin');
        // Background color cactus
        addBgColor(cell, colors.cactus);
    } else {
        // Background color silver
        addBgColor(cell, colors.silver);
    }

    // Border bottom medium
    addBorder(cell, 'bottom');
});

const thirteenAndFourteen = (cell, idx) => {
    switch (idx) {
        case 0:
            // border bottom hair font bold
            addBorder(cell, 'bottom', 'hair');
            addFont(cell, 'bold', true);
            break;
        case 1:
            // border left bottom thin background color cactus
            addBorder(cell, 'left', 'thin');
            addBorder(cell, 'bottom', 'thin');
            addBgColor(cell, colors.cactus);
            break;
        case 2:
            // border bottom thin
            addBorder(cell, 'bottom', 'thin');
            break;
        case 3:
            // border bottom right thing
            addBorder(cell, 'bottom', 'thin');
            addBorder(cell, 'right', 'thin');
            break;
        default:
            break;
    }
};
// Cells A13 - D13
const cellThirteen = ['A13', 'B13', 'C13', 'D13'];
// Cells A14 - D14
const cellFourteen = ['A14', 'B14', 'C14', 'D14'];
// Call function thirteenAndFourteen on variable cellThirteen and cellFourteen
cellThirteen.forEach(thirteenAndFourteen);
cellFourteen.forEach(thirteenAndFourteen);

// Cell A13 font bold
addFont('A13', 'bold', true);
// Cell B13 background color cactus and border bottom light
addBgColor('B13', colors.cactus);

// Cell A14 font bold
addFont('A14', 'bold', true);
// Cell B14 background color cactus and border bottom light
addBgColor('B14', colors.cactus);
addBorder('B14', 'bottom', 'thin');

// Cells A15 - Q15
const cellsFifteen = ['A15', 'B15', 'C15', 'D15', 'E15', 'F15', 'G15', 'H15', 'I15', 'J15', 'K15', 'L15', 'M15', 'N15', 'O15', 'P15', 'Q15'];
cellsFifteen.forEach((cell, idx) => {
    if (idx === 0) {
        // Font bold
        addFont(cell, 'bold', true);
        // Background color baby blue
        addBgColor(cell, colors.lightBlue);
    } else if (idx != 1) {
        addBgColor(cell, colors.lightBlue);
    }
    // Border top
    addBorder(cell, 'top');
});

// Cells A16 - Q16
const cellsSixteen = ['A16', 'B16', 'C16', 'D16', 'E16', 'F16', 'G16', 'H16', 'I16', 'J16', 'K16', 'L16', 'M16', 'N16', 'O16', 'P16', 'Q16'];
cellsSixteen.forEach((cell, idx) => {
    if (idx === 0) {
        addBgColor(cell, colors.lightBlue);
    } else if (idx % 2 !== 0) {
        // If idx is odd, give border left, center text size 8 bold
        addBorder(cell, 'left', 'thin');
        addBorder(cell, 'bottom', 'thin');
        addBorder(cell, 'top', 'thin');
        addAlignment(cell, 'horizontal', 'center');
        addFont(cell, 'sz', 8);
        addFont(cell, 'bold', true);
        addBgColor(cell, colors.babyBlue);
    } else {
        // If idx is even number, excluding 0, add light border right and bottom
        addBorder(cell, 'right', 'thin');
        addBorder(cell, 'bottom', 'thin');
        addBorder(cell, 'top', 'thin');
    }
});

// Cells A17 - Q17
const cellsSeventeen = ['A17', 'B17', 'C17', 'D17', 'E17', 'F17', 'G17', 'H17', 'I17', 'J17', 'K17', 'L17', 'M17', 'N17', 'O17', 'P17', 'Q17'];
cellsSeventeen.forEach((cell, idx) => {
    if (idx === 0) {
        addBgColor(cell, colors.lightBlue);
    } else {
        // Border left right bottom thin and center text size 8 bold
        addBorder(cell, 'left', 'thin');
        addBorder(cell, 'right', 'thin');
        addBorder(cell, 'bottom', 'thin');
        addAlignment(cell, 'horizontal', 'center');
        addFont(cell, 'sz', 8);
        addFont(cell, 'bold', true);
        addBgColor(cell, colors.babyBlue);
    }
});

// Cells A18 - Q18
const cellsEighteen = ['A18', 'B18', 'C18', 'D18', 'E18', 'F18', 'G18', 'H18', 'I18', 'J18', 'K18', 'L18', 'M18', 'N18', 'O18', 'P18', 'Q18'];
cellsEighteen.forEach((cell, idx) => {
    if (idx === 0) {
        addBgColor(cell, colors.lightBlue);
    } else {
        // Border left right bottom light and background color cactus
        addBorder(cell, 'left', 'thin');
        addBorder(cell, 'right', 'thin');
        addBorder(cell, 'bottom', 'thin');
        addBgColor(cell, colors.cactus);
    }
});

// Cells A19 - Q19
const cellsNineteen = ['A19', 'B19', 'C19', 'D19', 'E19', 'F19', 'G19', 'H19', 'I19', 'J19', 'K19', 'L19', 'M19', 'N19', 'O19', 'P19', 'Q19'];
cellsNineteen.forEach((cell, idx) => {
    if (idx === 0) {
        addBgColor(cell, colors.lightBlue);
    } else {
        // Border left right bottom light text size 8 bold align center
        addBorder(cell, 'left', 'thin');
        addBorder(cell, 'right', 'thin');
        addBorder(cell, 'bottom', 'thin');
        addAlignment(cell, 'horizontal', 'center');
        addFont(cell, 'sz', 8);
        addFont(cell, 'bold', true);
        addBgColor(cell, colors.babyBlue);
    }
});

// Cells A20 - Q20
const cellsTwenty = ['A20', 'B20', 'C20', 'D20', 'E20', 'F20', 'G20', 'H20', 'I20', 'J20', 'K20', 'L20', 'M20', 'N20', 'O20', 'P20', 'Q20'];
cellsTwenty.forEach((cell, idx) => {
    if (idx === 0) {
        addBgColor(cell, colors.lightBlue);
        // Text align right bold size 10
        addAlignment(cell, 'horizontal', 'right');
        addFont(cell, 'bold', true);
        addFont(cell, 'sz', 10);
    } else {
        // Border left right light background color cactus
        addBorder(cell, 'left', 'thin');
        addBorder(cell, 'right', 'thin');
        addBgColor(cell, colors.cactus);
    }
    // Border bottom
    addBorder(cell, 'bottom');
});

// Cells A21 - Q21
const cellsTwentyOne = ['A21', 'B21', 'C21', 'D21', 'E21', 'F21', 'G21', 'H21', 'I21', 'J21', 'K21', 'L21', 'M21', 'N21', 'O21', 'P21', 'Q21'];
cellsTwentyOne.forEach((cell, idx) => {
    if (idx === 0) {
        // Background color light tan bold text
        addBgColor(cell, colors.lightTan);
        addFont(cell, 'bold', true);
    } else if (idx % 2 !== 0) {
        // if idx is odd, border left bottom light background color tan center text bold size 8
        addBorder(cell, 'left', 'thin');
        addBorder(cell, 'bottom', 'thin');
        addBgColor(cell, colors.tan);
        addAlignment(cell, 'horizontal', 'center');
        addFont(cell, 'bold', true);
        addFont(cell, 'sz', 8);
    } else {
        // if idx is even, border right bottom light
        addBorder(cell, 'right', 'thin');
        addBorder(cell, 'bottom', 'thin');
    }
});

// Cells A22 - Q22
const cellsTwentyTwo = ['A22', 'B22', 'C22', 'D22', 'E22', 'F22', 'G22', 'H22', 'I22', 'J22', 'K22', 'L22', 'M22', 'N22', 'O22', 'P22', 'Q22'];
cellsTwentyTwo.forEach((cell, idx) => {
    if (idx === 0) {
        // Background color light tan
        addBgColor(cell, colors.lightTan);
    } else {
        // Border left bottom right light center bold size 8
        addBorder(cell, 'left', 'thin');
        addBorder(cell, 'bottom', 'thin');
        addBorder(cell, 'right', 'thin');
        addAlignment(cell, 'horizontal', 'center');
        addFont(cell, 'bold', true);
        addFont(cell, 'sz', 8);
    }
});

// Cells A23 - Q23
const cellsTwentyThree = ['A23', 'B23', 'C23', 'D23', 'E23', 'F23', 'G23', 'H23', 'I23', 'J23', 'K23', 'L23', 'M23', 'N23', 'O23', 'P23', 'Q23'];

cellsTwentyThree.forEach((cell, idx) => {
    if (idx === 0) {
        // Background color light tan
        addBgColor(cell, colors.lightTan);
    } else {
        // Border left bottom right background color cactus
        addBorder(cell, 'left', 'thin');
        addBorder(cell, 'bottom', 'thin');
        addBorder(cell, 'right', 'thin');
        addBgColor(cell, colors.cactus);
    }
});

// Cells A24 - I24
const cellsTwentyFour = ['A24', 'B24', 'C24', 'D24', 'E24', 'F24', 'G24', 'H24', 'I24'];

cellsTwentyFour.forEach((cell, idx) => {
    if (idx === 0) {
        // Background color light tan
        addBgColor(cell, colors.lightTan);
    } else if (idx % 2 !== 0) {
        // if idx is odd, border left bottom light background color tan center text bold size 8
        addBorder(cell, 'left', 'thin');
        addBorder(cell, 'bottom', 'thin');
        addBgColor(cell, colors.tan);
        addAlignment(cell, 'horizontal', 'center');
        addFont(cell, 'bold', true);
        addFont(cell, 'sz', 8);
    } else {
        // if idx is even, border right bottom light
        addBorder(cell, 'right', 'thin');
        addBorder(cell, 'bottom', 'thin');
    }
});

// Cells A25 - I25
const cellsTwentyFive = ['A25', 'B25', 'C25', 'D25', 'E25', 'F25', 'G25', 'H25', 'I25'];

cellsTwentyFive.forEach((cell, idx) => {
    if (idx === 0) {
        // Background color light tan
        addBgColor(cell, colors.lightTan);
    } else {
        // Border left bottom right light center bold size 8
        addBorder(cell, 'left', 'thin');
        addBorder(cell, 'bottom', 'thin');
        addBorder(cell, 'right', 'thin');
        addAlignment(cell, 'horizontal', 'center');
        addFont(cell, 'bold', true);
        addFont(cell, 'sz', 8);
    }
});

// Cells A26 - I26
const cellsTwentySix = ['A26', 'B26', 'C26', 'D26', 'E26', 'F26', 'G26', 'H26', 'I26'];

cellsTwentySix.forEach((cell, idx) => {
    if (idx === 0) {
        // Background color light tan
        addBgColor(cell, colors.lightTan);
    } else {
        // Border left bottom right background color cactus
        addBorder(cell, 'left', 'thin');
        addBorder(cell, 'bottom', 'thin');
        addBorder(cell, 'right', 'thin');
        addBgColor(cell, colors.cactus);
    }
});

// Loop through rows 24 - 26 and cols J - Q
for (let i = 24; i <= 26; i++) {
    for (let j = 10; j <= 17; j++) {
        let cell = String.fromCharCode(64 + j) + i;
        // Background color light tan
        addBgColor(cell, colors.lightTan);
    }
}

// Cells A27 - Q27
const cellsTwentySeven = ['A27', 'B27', 'C27', 'D27', 'E27', 'F27', 'G27', 'H27', 'I27', 'J27', 'K27', 'L27', 'M27', 'N27', 'O27', 'P27', 'Q27'];

cellsTwentySeven.forEach((cell, idx) => {
    if (idx == 0) {
        // Bold font size 10
        addFont(cell, 'bold', true);
        addFont(cell, 'sz', 10);
    } else if (idx == 1) {
        // Background color cactus border left thin
        addBgColor(cell, colors.cactus);
        addBorder(cell, 'left', 'thin');
    } else if (idx == 3) {
        // Border right thin
        addBorder(cell, 'right', 'thin');
    }
    // Border bottom top
    addBorder(cell, 'bottom');
    addBorder(cell, 'top');
});

// Cells A28 - O28
const cellsTwentyEight = ['A28', 'B28', 'C28', 'D28', 'E28', 'F28', 'G28', 'H28', 'I28', 'J28', 'K28', 'L28', 'M28', 'N28', 'O28'];

cellsTwentyEight.forEach((cell, idx) => {
    if (idx == 0) {
        // Background color rain bold
        addBgColor(cell, colors.rain);
        addFont(cell, 'bold', true);
    } else if (idx % 2 != 0) {
        // if idx is odd, border thin left bottom background color cloudy center text bold size 8
        addBorder(cell, 'left', 'thin');
        addBorder(cell, 'bottom', 'thin');
        addBgColor(cell, colors.cloudy);
        addAlignment(cell, 'horizontal', 'center');
        addFont(cell, 'bold', true);
        addFont(cell, 'sz', 8);
    } else {
        // if idx is even, border thin right bottom
        addBorder(cell, 'right', 'thin');
        addBorder(cell, 'bottom', 'thin');
    }
});

// Cells A29 - O29
const cellsTwentyNine = ['A29', 'B29', 'C29', 'D29', 'E29', 'F29', 'G29', 'H29', 'I29', 'J29', 'K29', 'L29', 'M29', 'N29', 'O29'];

cellsTwentyNine.forEach((cell, idx) => {
    if (idx == 0) {
        // Background color rain
        addBgColor(cell, colors.rain);
    } else {
        // Border left right bottom thin center bold size 8 background color cloudy
        addBorder(cell, 'left', 'thin');
        addBorder(cell, 'right', 'thin');
        addBorder(cell, 'bottom', 'thin');
        addAlignment(cell, 'horizontal', 'center');
        addFont(cell, 'bold', true);
        addFont(cell, 'sz', 8);
        addBgColor(cell, colors.cloudy);
    }
});

// Cells A30 - O30
const cellsThirty = ['A30', 'B30', 'C30', 'D30', 'E30', 'F30', 'G30', 'H30', 'I30', 'J30', 'K30', 'L30', 'M30', 'N30', 'O30'];

cellsThirty.forEach((cell, idx) => {
    if (idx == 0) {
        // Background color rain bold right align
        addBgColor(cell, colors.rain);
        addFont(cell, 'bold', true);
        addAlignment(cell, 'horizontal', 'right');
    } else {
        // Border left right bottom thin background color cactus
        addBorder(cell, 'left', 'thin');
        addBorder(cell, 'right', 'thin');
        addBorder(cell, 'bottom', 'thin');
        addBgColor(cell, colors.cactus);
    }
});

// Cells A31 - O31
const cellsThirtyOne = ['A31', 'B31', 'C31', 'D31', 'E31', 'F31', 'G31', 'H31', 'I31', 'J31', 'K31', 'L31', 'M31', 'N31', 'O31'];

cellsThirtyOne.forEach((cell, idx) => {
    if (idx == 0) {
        // Background color rain
        addBgColor(cell, colors.rain);
    } else if (idx % 2 != 0) {
        // If idx is odd, center text bold size 8 background color cloudy border left bottom right thin
        addAlignment(cell, 'horizontal', 'center');
        addFont(cell, 'bold', true);
        addFont(cell, 'sz', 8);
        addBgColor(cell, colors.cloudy);
        addBorder(cell, 'left', 'thin');
        addBorder(cell, 'bottom', 'thin');
    } else {
        // If idx is even, border right bottom thin
        addBorder(cell, 'right', 'thin');
        addBorder(cell, 'bottom', 'thin');
    }
});

// Cells A32 - O32
const cellsThirtyTwo = ['A32', 'B32', 'C32', 'D32', 'E32', 'F32', 'G32', 'H32', 'I32', 'J32', 'K32', 'L32', 'M32', 'N32', 'O32'];

cellsThirtyTwo.forEach((cell, idx) => {
    if (idx == 0) {
        // Background color rain
        addBgColor(cell, colors.rain);
    } else {
        // Border left right bottom thing background color cloudy center text bold size 8
        addBorder(cell, 'left', 'thin');
        addBorder(cell, 'right', 'thin');
        addBorder(cell, 'bottom', 'thin');
        addBgColor(cell, colors.cloudy);
        addAlignment(cell, 'horizontal', 'center');
        addFont(cell, 'bold', true);
        addFont(cell, 'sz', 8);
    }
});

// Cells A33 - O33
const cellsThirtyThree = ['A33', 'B33', 'C33', 'D33', 'E33', 'F33', 'G33', 'H33', 'I33', 'J33', 'K33', 'L33', 'M33', 'N33', 'O33'];

cellsThirtyThree.forEach((cell, idx) => {
    if (idx == 0) {
        // Background color rain bold right align
        addBgColor(cell, colors.rain);
        addFont(cell, 'bold', true);
        addAlignment(cell, 'horizontal', 'right');
    } else {
        // Border left right bottom thin background color cactus
        addBorder(cell, 'left', 'thin');
        addBorder(cell, 'right', 'thin');
        addBorder(cell, 'bottom', 'thin');
        addBgColor(cell, colors.cactus);
    }
});

// Cells A34 - O34
const cellsThirtyFour = ['A34', 'B34', 'C34', 'D34', 'E34', 'F34', 'G34', 'H34', 'I34', 'J34', 'K34', 'L34', 'M34', 'N34', 'O34'];

cellsThirtyFour.forEach((cell, idx) => {
    if (idx == 0) {
        // Background color rain bold right align
        addBgColor(cell, colors.rain);
        addFont(cell, 'bold', true);
        addAlignment(cell, 'horizontal', 'right');
    } else if (idx == 1) {
        // Border left bottom thin background color cactus
        addBorder(cell, 'left', 'thin');
        addBorder(cell, 'bottom', 'thin');
        addBgColor(cell, colors.cactus);
    } else  if (idx == 2) {
        // Border left bottom thin background color cactus
        addBorder(cell, 'left', 'thin');
        addBorder(cell, 'bottom', 'thin');
        addBgColor(cell, colors.cactus);
    } else  if (idx == 3) {
        // Border right bottom thin background color cactus
        addBorder(cell, 'right', 'thin');
        addBorder(cell, 'bottom', 'thin');
        addBgColor(cell, colors.cactus);
    } else if (idx == 4) {
        // Center bold background color rain bottom left thin
        addAlignment(cell, 'horizontal', 'center');
        addFont(cell, 'bold', true);
        addBgColor(cell, colors.rain);
        addBorder(cell, 'left', 'thin');
        addBorder(cell, 'bottom', 'thin');
    } else if (idx == 5) {
        // border right bottom thin
        addBorder(cell, 'right', 'thin');
        addBorder(cell, 'bottom', 'thin');
    } else if (idx == 6) {
        // border left bottom thin background color cactus
        addBorder(cell, 'left', 'thin');
        addBorder(cell, 'bottom', 'thin');
        addBgColor(cell, colors.cactus);
    } else if (idx < cellsThirtyFour.length - 1) {
        // border bottom cactus
        addBorder(cell, 'bottom', 'thin');
        addBgColor(cell, colors.cactus);
    } else {
        // border right bottom thin background color cactus
        addBorder(cell, 'right', 'thin');
        addBorder(cell, 'bottom', 'thin');
        addBgColor(cell, colors.cactus);
    }
});

// Cells thirty four + cells P34 - O34
const cellsThirtyFourRow = [...cellsThirtyFour, 'P34', 'Q34'];
cellsThirtyFourRow.forEach((cell, idx) => {
    // Cell bottom border
    addBorder(cell, 'bottom');
});

// Loop through rows 28 - 34 and cols P - Q
for (let i = 28; i <= 34; i++) {
    for (let j = 16; j <= 17; j++) {
        let cell = String.fromCharCode(64 + j) + i;
        // Background color rain
        addBgColor(cell, colors.rain);
    }
}

// Cells A35 - H35
const cellsThirtyFive = ['A35', 'B35', 'C35', 'D35', 'E35', 'F35', 'G35', 'H35'];

cellsThirtyFive.forEach((cell, idx) => {
    if (idx == 0) {
        // bold left align size 10
        addFont(cell, 'bold', true);
        addAlignment(cell, 'horizontal', 'left');
        addFont(cell, 'sz', 10);
    } else if (idx == 1 || idx == 6) {
        // border left thin background color cactus
        addBorder(cell, 'left', 'thin');
        addBgColor(cell, colors.cactus);
    } else if (idx == 3 || idx == 7) {
        // border right thin
        addBorder(cell, 'right', 'thin');
    } else if (idx == 4) {
        // center bold size 10
        addAlignment(cell, 'horizontal', 'center');
        addFont(cell, 'bold', true);
        addFont(cell, 'sz', 10);
    }
});

// Cells A36 - Q36
const cellsThirtySix = ['A36', 'B36', 'C36', 'D36', 'E36', 'F36', 'G36', 'H36', 'I36', 'J36', 'K36', 'L36', 'M36', 'N36', 'O36', 'P36', 'Q36'];

cellsThirtySix.forEach((cell, idx) => {
    if (idx == 0) {
        // bold left align
        addFont(cell, 'bold', true);
        addAlignment(cell, 'horizontal', 'left');
    }
    // border top background color rain
    addBorder(cell, 'top');
    addBgColor(cell, colors.rain);
});

// Cells A37 - Q37
const cellsThirtySeven = ['A37', 'B37', 'C37', 'D37', 'E37', 'F37', 'G37', 'H37', 'I37', 'J37', 'K37', 'L37', 'M37', 'N37', 'O37', 'P37', 'Q37'];

cellsThirtySeven.forEach((cell, idx) => {
    if (idx == 0 || idx == 1) {
        // background color rain
        addBgColor(cell, colors.rain);
    } else if (idx == 2 || idx == 5 || idx == 7 || idx == 9 || idx == 11 || idx == 13 || idx == 15) {
        // border left thin cloudy center text bold size 9
        addBorder(cell, 'left', 'thin');
        addBgColor(cell, colors.cloudy);
        addAlignment(cell, 'horizontal', 'center');
        addFont(cell, 'bold', true);
        addFont(cell, 'sz', 9);
    } else if (idx != 3) {
        // border right thin
        addBorder(cell, 'right', 'thin');
    }

    if (idx > 1) {
        addBorder(cell, 'bottom', 'thin');
        addBorder(cell, 'top', 'thin');
    }
});

// Cells A38 - Q38
const cellsThirtyEight = ['A38', 'B38', 'C38', 'D38', 'E38', 'F38', 'G38', 'H38', 'I38', 'J38', 'K38', 'L38', 'M38', 'N38', 'O38', 'P38', 'Q38'];

cellsThirtyEight.forEach((cell, idx) => {
    if (idx == 0) {
        // bold right align bg color rain
        addFont(cell, 'bold', true);
        addAlignment(cell, 'horizontal', 'right');
        addBgColor(cell, colors.rain);
    } else if (idx == 1) {
        // border left thing cactus
        addBorder(cell, 'left', 'thin');
        addBgColor(cell, colors.cactus);
    } else if (idx == 2 || idx == 4) {
        // border right thin
        addBorder(cell, 'right', 'thin');
    } else if (idx == 3) {
        // border left thin cloudy center text bold
        addBorder(cell, 'left', 'thin');
        addBgColor(cell, colors.cloudy);
        addAlignment(cell, 'horizontal', 'center');
        addFont(cell, 'bold', true);
    } else {
        // border left right thin bold center sz 8
        addBorder(cell, 'left', 'thin');
        addBorder(cell, 'right', 'thin');
        addFont(cell, 'bold', true);
        addAlignment(cell, 'horizontal', 'center');
        addFont(cell, 'sz', 8);
        addBgColor(cell, colors.cloudy);
    }

    if (idx != 0) {
        // border top bottom thin
        addBorder(cell, 'top', 'thin');
        addBorder(cell, 'bottom', 'thin');
    }
});

// Cells A39 - Q39
const cellsThirtyNine = ['A39', 'B39', 'C39', 'D39', 'E39', 'F39', 'G39', 'H39', 'I39', 'J39', 'K39', 'L39', 'M39', 'N39', 'O39', 'P39', 'Q39'];

cellsThirtyNine.forEach((cell, idx) => {
    if (idx == 0) {
        // bold right align bg color rain
        addFont(cell, 'bold', true);
        addAlignment(cell, 'horizontal', 'right');
        addBgColor(cell, colors.rain);
    } else if (idx == 1) {
        // border left thing cactus
        addBorder(cell, 'left', 'thin');
        addBgColor(cell, colors.cactus);
    } else if (idx == 2 || idx == 4) {
        if (idx == 4) {
            // background color rain
            addBgColor(cell, colors.rain);
        }
        // border right thin
        addBorder(cell, 'right', 'thin');
    } else if (idx == 3) {
        // border left thin rain
        addBorder(cell, 'left', 'thin');
        addBgColor(cell, colors.rain);
    }  else {
        // border left right thin bg color cactus
        addBorder(cell, 'left', 'thin');
        addBorder(cell, 'right', 'thin');
        addBgColor(cell, colors.cactus);
    }

    // for cells I39, K39, M39, O39, Q39
    if (idx % 2 == 0 && idx >= 8) {
        // bg pink
        addBgColor(cell, colors.pink);
    }

    if (idx != 0) {
        // border top bottom thin
        addBorder(cell, 'top', 'thin');
        addBorder(cell, 'bottom', 'thin');
    }
});

// Cells A40 - Q40
const cellsForty = ['A40', 'B40', 'C40', 'D40', 'E40', 'F40', 'G40', 'H40', 'I40', 'J40', 'K40', 'L40', 'M40', 'N40', 'O40', 'P40', 'Q40'];

cellsForty.forEach((cell, idx) => {
    if (idx == 0) {
        // bold right align bg color rain
        addFont(cell, 'bold', true);
        addAlignment(cell, 'horizontal', 'right');
        addBgColor(cell, colors.rain);
    } else if (idx == 3) {
        // border left thin cloudy center text bold
        addBorder(cell, 'left', 'thin');
        addBgColor(cell, colors.cloudy);
        addAlignment(cell, 'horizontal', 'center');
        addFont(cell, 'bold', true);
    } else if (idx % 2 == 1) {
        // border left thin cactus
        addBorder(cell, 'left', 'thin');
        addBgColor(cell, colors.cactus);
    } else {
        // border right thin
        addBorder(cell, 'right', 'thin');
    }

    if (idx != 0) {
        // border top bottom thin
        addBorder(cell, 'top', 'thin');
        addBorder(cell, 'bottom', 'thin');
    }
});

// Cells A41 - Q41
const cellsFortyOne = ['A41', 'B41', 'C41', 'D41', 'E41', 'F41', 'G41', 'H41', 'I41', 'J41', 'K41', 'L41', 'M41', 'N41', 'O41', 'P41', 'Q41'];

cellsFortyOne.forEach((cell, idx) => {
    if (idx < 5) {
        // bg color rain
        addBgColor(cell, colors.rain);
    } else if (idx % 2 == 1) {
        // border top bottom left thin cloudy center text bold size 8
        addBorder(cell, 'left', 'thin');
        addBorder(cell, 'top', 'thin');
        addBorder(cell, 'bottom', 'thin');
        addBgColor(cell, colors.cloudy);
        addAlignment(cell, 'horizontal', 'center');
        addFont(cell, 'bold', true);
        addFont(cell, 'sz', 8);
    } else {
        // border right top bottom thin
        addBorder(cell, 'right', 'thin');
        addBorder(cell, 'top', 'thin');
        addBorder(cell, 'bottom', 'thin');
    }
});

// Cells A42 - Q42
const cellsFortyTwo = ['A42', 'B42', 'C42', 'D42', 'E42', 'F42', 'G42', 'H42', 'I42', 'J42', 'K42', 'L42', 'M42', 'N42', 'O42', 'P42', 'Q42'];

cellsFortyTwo.forEach((cell, idx) => {
    if (idx < 3) {
        // bg color rain
        addBgColor(cell, colors.rain);
    } else if (idx == 3) {
        // border left thin cloudy center text bold
        addBorder(cell, 'left', 'thin');
        addBgColor(cell, colors.cloudy);
        addAlignment(cell, 'horizontal', 'center');
        addFont(cell, 'bold', true);
    } else if (idx == 4) {
        // border right thin 
        addBorder(cell, 'right', 'thin');
    } else {
        // border left right thin center text bold size 8
        addBorder(cell, 'left', 'thin');
        addBorder(cell, 'right', 'thin');
        addAlignment(cell, 'horizontal', 'center');
        addFont(cell, 'bold', true);
        addFont(cell, 'sz', 8);
    }

    if (idx >= 3) {
        // border top bottom thin
        addBorder(cell, 'top', 'thin');
        addBorder(cell, 'bottom', 'thin');
    }
});

// Cells A43 - Q43
const cellsFortyThree = ['A43', 'B43', 'C43', 'D43', 'E43', 'F43', 'G43', 'H43', 'I43', 'J43', 'K43', 'L43', 'M43', 'N43', 'O43', 'P43', 'Q43'];

cellsFortyThree.forEach((cell, idx) => {
    if (idx < 5) {
        // bg color rain
        addBgColor(cell, colors.rain);
    } else {
        // border left right bottom top thin
        addBorder(cell, 'left', 'thin');
        addBorder(cell, 'right', 'thin');
        addBorder(cell, 'bottom', 'thin');
        addBorder(cell, 'top', 'thin');
        if (idx >= 8 && idx % 2 == 0) {
            // bg color pink
            addBgColor(cell, colors.pink);
        } else {
            // bg color cactus
            addBgColor(cell, colors.cactus);
        }
    }
});

// Cells A44 - Q44
const cellsFortyFour = ['A44', 'B44', 'C44', 'D44', 'E44', 'F44', 'G44', 'H44', 'I44', 'J44', 'K44', 'L44', 'M44', 'N44', 'O44', 'P44', 'Q44'];

cellsFortyFour.forEach((cell, idx) => {
    if (idx < 3) {
        // bg color rain
        addBgColor(cell, colors.rain);
    } else if (idx % 2 == 1) {
        // border left top bottom thin
        addBorder(cell, 'left', 'thin');
        addBorder(cell, 'top', 'thin');
        addBorder(cell, 'bottom', 'thin');
        if (idx == 3) {
            // center text bold bg color cloudy
            addAlignment(cell, 'horizontal', 'center');
            addFont(cell, 'bold', true);
            addBgColor(cell, colors.cloudy);
        } else {
            // bg color cactus
            addBgColor(cell, colors.cactus);
        }
    } else {
        // border right top bottom thin
        addBorder(cell, 'right', 'thin');
        addBorder(cell, 'top', 'thin');
        addBorder(cell, 'bottom', 'thin');
    }
});

// Cells A45 - Q45
const cellsFortyFive = ['A45', 'B45', 'C45', 'D45', 'E45', 'F45', 'G45', 'H45', 'I45', 'J45', 'K45', 'L45', 'M45', 'N45', 'O45', 'P45', 'Q45'];

cellsFortyFive.forEach((cell, idx) => {
    if (idx == 0) {
        // bg color rain
        addBgColor(cell, colors.rain);
    } else if (idx == 1) {
        // border left thin background color cloudy center text bold
        addBorder(cell, 'left', 'thin');
        addBgColor(cell, colors.cloudy);
        addAlignment(cell, 'horizontal', 'center');
        addFont(cell, 'bold', true);
    } else if (idx >= 4 && idx % 2 == 0) {
        // border right thin
        addBorder(cell, 'right', 'thin');
    } else if (idx > 4) {
        // border left thin bg color sky blue center text bold
        addBorder(cell, 'left', 'thin');
        addBgColor(cell, colors.skyBlue);
        addAlignment(cell, 'horizontal', 'center');
        addFont(cell, 'bold', true);
    }

    if (idx > 0) {
        // border top bottom thin
        addBorder(cell, 'top', 'thin');
        addBorder(cell, 'bottom', 'thin');
    }
});

// Cells A46 - Q46
const cellsFortySix = ['A46', 'B46', 'C46', 'D46', 'E46', 'F46', 'G46', 'H46', 'I46', 'J46', 'K46', 'L46', 'M46', 'N46', 'O46', 'P46', 'Q46'];

cellsFortySix.forEach((cell, idx) => {
    if (idx == 0) {
        // bg color rain right align bold
        addBgColor(cell, colors.rain);
        addAlignment(cell, 'horizontal', 'right');
        addFont(cell, 'bold', true);
    } else if (idx == 1 || idx == 4 || idx == 7 || idx == 13) {
        // border left thin bg color cloudy center text bold
        addBorder(cell, 'left', 'thin');
        addBgColor(cell, colors.cloudy);
        addAlignment(cell, 'horizontal', 'center');
        addFont(cell, 'bold', true);
    } else if (idx == 2 || idx == 5 || idx == 8 || idx == 14) {
        // borer right thin
        addBorder(cell, 'right', 'thin');
    } else if (idx == 10) {
        // border left right bg color cloudy bold
        addBorder(cell, 'left', 'thin');
        addBorder(cell, 'right', 'thin');
        addBgColor(cell, colors.cloudy);
        addFont(cell, 'bold', true);
    } else if (idx < 10) {
        // border left right thin bg color cactus
        addBorder(cell, 'left', 'thin');
        addBorder(cell, 'right', 'thin');
        addBgColor(cell, colors.cactus);
    } else if (idx == 11 || idx == 15) {
        // border left thin bg color cactus
        addBorder(cell, 'left', 'thin');
        addBgColor(cell, colors.cactus);
    } else if (idx == 12 || idx == 16) {
        // border right thin bg color cactus
        addBorder(cell, 'right', 'thin');
        addBgColor(cell, colors.cactus);
    }

    // border bottom
    addBorder(cell, 'bottom');
});

// Cell A47 bold left align
addFont('A47', 'bold', true);
addAlignment('A47', 'horizontal', 'left');

// Cells D48 - Q48
const cellsFortyEight = ['D48', 'E48', 'F48', 'G48', 'H48', 'I48', 'J48', 'K48', 'L48', 'M48', 'N48', 'O48', 'P48', 'Q48'];

cellsFortyEight.forEach((cell, idx) => {
    if (idx % 2 == 0) {
        // border left thin bold center sz 9
        addBorder(cell, 'left', 'thin');
        addFont(cell, 'bold', true);
        addAlignment(cell, 'horizontal', 'center');
        addFont(cell, 'sz', 9);
    } else {
        // border right thin
        addBorder(cell, 'right', 'thin');
    }
    // border top bottom thin
    addBorder(cell, 'top', 'thin');
    addBorder(cell, 'bottom', 'thin');
});

// Loop through rows 49 - 51 and columns A - Q
for (let i = 49; i <= 51; i++) {
    for (let j = 0; j <= 16; j++) {
        const cell = String.fromCharCode(65 + j) + i;
        if (j == 0 && i == 49) {
            // bold right align
            addFont(cell, 'bold', true);
            addAlignment(cell, 'horizontal', 'right');
        } else if (j == 1) {
            // border left thin bg color cactus
            addBorder(cell, 'left', 'thin');
            addBgColor(cell, colors.cactus);
        } else if (j == 4) {
            // border right thin
            addBorder(cell, 'right', 'thin');
        } else if (j > 4 && j % 2 == 1) {
            // border left thin cactus
            addBorder(cell, 'left', 'thin');
            addBgColor(cell, colors.cactus);
        } else if (j > 4 && j % 2 == 0) {
            // border right thin
            addBorder(cell, 'right', 'thin');
        }
        if (j != 0) {
            // border top bottom thin
            addBorder(cell, 'top', 'thin');
            addBorder(cell, 'bottom', 'thin');
        }
    }
}

// Cells D52 - Q52
const cellsFiftyTwo = ['D52', 'E52', 'F52', 'G52', 'H52', 'I52', 'J52', 'K52', 'L52', 'M52', 'N52', 'O52', 'P52', 'Q52'];

cellsFiftyTwo.forEach((cell, idx) => {
    if (idx % 2 == 0) {
        // border left thin bold center sz 9
        addBorder(cell, 'left', 'thin');
        addFont(cell, 'bold', true);
        addAlignment(cell, 'horizontal', 'center');
        addFont(cell, 'sz', 9);
    } else {
        // border right thin
        addBorder(cell, 'right', 'thin');
    }
    // border top bottom thin
    addBorder(cell, 'top', 'thin');
    addBorder(cell, 'bottom', 'thin');
});

// Loop through rows 53 - 54 and columns A - Q
for (let i = 53; i <= 54; i++) {
    for (let j = 0; j <= 16; j++) {
        const cell = String.fromCharCode(65 + j) + i;
        if (j == 0 && i == 53) {
            // bold right align
            addFont(cell, 'bold', true);
            addAlignment(cell, 'horizontal', 'right');
        } else if (j == 1) {
            // border left thin bg color cactus
            addBorder(cell, 'left', 'thin');
            addBgColor(cell, colors.cactus);
        } else if (j == 4) {
            // border right thin
            addBorder(cell, 'right', 'thin');
        } else if (j > 4 && j % 2 == 1) {
            // border left thin cactus
            addBorder(cell, 'left', 'thin');
            addBgColor(cell, colors.cactus);
        } else if (j > 4 && j % 2 == 0) {
            // border right thin
            addBorder(cell, 'right', 'thin');
        }
        if (j != 0) {
            // border top bottom thin
            addBorder(cell, 'top', 'thin');
            addBorder(cell, 'bottom', 'thin');
        }
    }
}

// Cells A55 - Q55
const cellsFiftyFive = ['A55', 'B55', 'C55', 'D55', 'E55', 'F55', 'G55', 'H55', 'I55', 'J55', 'K55', 'L55', 'M55', 'N55', 'O55', 'P55', 'Q55'];

cellsFiftyFive.forEach((cell, idx) => {
    if (idx == 0) {
        // right align bold
        addAlignment(cell, 'horizontal', 'right');
        addFont(cell, 'bold', true);
    } else if (idx == 1 || idx == 6) {
        // border left cactus
        addBorder(cell, 'left', 'thin');
        addBgColor(cell, colors.cactus);
    } else if (idx == 3 || idx == 16) {
        // border right thin
        addBorder(cell, 'right', 'thin');
    } else if (idx == 4) {
        // center text bold
        addAlignment(cell, 'horizontal', 'center');
        addFont(cell, 'bold', true);
    } 

    if (idx > 0) {
        // border top bottom thin
        addBorder(cell, 'top', 'thin');
        addBorder(cell, 'bottom', 'thin');
    }
});

// Cells A56 - Q56
const cellsFiftySix = ['A56', 'B56', 'C56', 'D56', 'E56', 'F56', 'G56', 'H56', 'I56', 'J56', 'K56', 'L56', 'M56', 'N56', 'O56', 'P56', 'Q56'];

cellsFiftySix.forEach((cell, idx) => {
    if (idx == 0) {
        // right align bold
        addAlignment(cell, 'horizontal', 'right');
        addFont(cell, 'bold', true);
    } else if (idx == 1) {
        // border left cactus
        addBorder(cell, 'left', 'thin');
        addBgColor(cell, colors.cactus);
    } else if (idx == 3) {
        // border right
        addBorder(cell, 'right', 'thin');
    }

    if (idx > 0) {
        // border top thin
        addBorder(cell, 'top', 'thin');
    }
    // border bottom
    addBorder(cell, 'bottom');
});

// Cells B57 - N57
const cellsFiftySeven = ['B57', 'C57', 'D57', 'E57', 'F57', 'G57', 'H57', 'I57', 'J57', 'K57', 'L57', 'M57', 'N57'];

cellsFiftySeven.forEach((cell, idx) => {
    if (idx < 8) {
        if (idx % 2 == 0) {
            // border left thin bold center
            addBorder(cell, 'left', 'thin');
            addFont(cell, 'bold', true);
            addAlignment(cell, 'horizontal', 'center');
        } else {
            // border right thin
            addBorder(cell, 'right', 'thin');
        }
    } else if (idx == 8 || idx == 9) {
        // border right left thin center bold
        addBorder(cell, 'right', 'thin');
        addBorder(cell, 'left', 'thin');
        addAlignment(cell, 'horizontal', 'center');
        addFont(cell, 'bold', true);
    } else if (idx == 10) {
        // border left thin center bold
        addBorder(cell, 'left', 'thin');
        addAlignment(cell, 'horizontal', 'center');
        addFont(cell, 'bold', true);
    } else if (idx == 12) {
        // border right thin
        addBorder(cell, 'right', 'thin');
    }

    // border top bottom thin
    addBorder(cell, 'top', 'thin');
    addBorder(cell, 'bottom', 'thin');
});

// Cells A58 - N58
const cellsFiftyEight = ['A58', 'B58', 'C58', 'D58', 'E58', 'F58', 'G58', 'H58', 'I58', 'J58', 'K58', 'L58', 'M58', 'N58'];

cellsFiftyEight.forEach((cell, idx) => {
    if (idx == 0) {
        // bold
        addFont(cell, 'bold', true);
    } else if (idx == 1 || idx == 3 || idx == 5) {
        // border left thin bg color cactus
        addBorder(cell, 'left', 'thin');
        addBgColor(cell, colors.cactus);
    } else if (idx == 2 || idx == 4 || idx == 6) {
        // border right thin
        addBorder(cell, 'right', 'thin');
    } else if (idx == 9 || idx == 10 || idx == 11) {
        // border left thin bg color cactus
        addBorder(cell, 'left', 'thin');
        addBgColor(cell, colors.cactus);
    } else if (idx == cellsFiftyEight.length - 1) {
        // border right thin
        addBorder(cell, 'right', 'thin');
    }
    
    if (idx > 0 && (idx < 7 || idx > 8)) {
        // border top bottom thin
        addBorder(cell, 'top', 'thin');
        addBorder(cell, 'bottom', 'thin');
    }
});

// Cells B59 - E59
const cellsFiftyNine = ['B59', 'C59', 'D59', 'E59'];

cellsFiftyNine.forEach((cell, idx) => {
    if (idx % 2 == 0) {
        // border left thin center bold
        addBorder(cell, 'left', 'thin');
        addAlignment(cell, 'horizontal', 'center');
        addFont(cell, 'bold', true);
    } else {
        // border right thin
        addBorder(cell, 'right', 'thin');
    }

    // border top bottom thin
    addBorder(cell, 'top', 'thin');
    addBorder(cell, 'bottom', 'thin');
});

// Cells A60 - E60
const cellsSixty = ['A60', 'B60', 'C60', 'D60', 'E60'];

cellsSixty.forEach((cell, idx) => {
    if (idx == 0) {
        // bold
        addFont(cell, 'bold', true);
    } else if (idx == 1 || idx == 3) {
        // border left thin bg color cactus
        addBorder(cell, 'left', 'thin');
        addBgColor(cell, colors.cactus);
    } else if (idx == 2 || idx == 4) {
        // border right thin
        addBorder(cell, 'right', 'thin');
    }

    if (idx > 0) {
        // border top bottom thin
        addBorder(cell, 'top', 'thin');
        addBorder(cell, 'bottom', 'thin');
    }
});

// Cells A61 - E61
const cellsSixtyOne = ['A61', 'B61', 'C61', 'D61', 'E61'];

cellsSixtyOne.forEach((cell, idx) => {
    if (idx == 0) {
        // bold
        addFont(cell, 'bold', true);
    } else if (idx == 1 || idx == 3) {
        // border left thin bg color cactus
        addBorder(cell, 'left', 'thin');
        addBgColor(cell, colors.cactus);
    } else if (idx == 2 || idx == 4) {
        // border right thin
        addBorder(cell, 'right', 'thin');
    }

    if (idx > 0) {
        // border top bottom thin
        addBorder(cell, 'top', 'thin');
        addBorder(cell, 'bottom', 'thin');
    }
});

// Cells A62 - D62
const cellsSixtyTwo = ['A62', 'B62', 'C62', 'D62'];

cellsSixtyTwo.forEach((cell, idx) => {
    if (idx == 0) {
        // bold
        addFont(cell, 'bold', true);
    } else if (idx == 1) {
        // border left thin bg color cactus
        addBorder(cell, 'left', 'thin');
        addBgColor(cell, colors.cactus);
    } else if (idx == 2) {
        // border right thin
        addBorder(cell, 'right', 'thin');
    } else {
        // bold size 12 dark blue
        addFont(cell, 'bold', true);
        addFont(cell, 'size', 12);
        addFont(cell, 'color', { rgb: colors.darkBlue });
    }
});

// Cells A63 - D63
const cellsSixtyThree = ['A63', 'B63', 'C63', 'D63'];

cellsSixtyThree.forEach((cell, idx) => {
    if (idx == 0) {
        // bold
        addFont(cell, 'bold', true);
    } else if (idx == 1) {
        // border left thin bg color cactus
        addBorder(cell, 'left', 'thin');
        addBgColor(cell, colors.cactus);
    } else if (idx == 2) {
        // border right thin
        addBorder(cell, 'right', 'thin');
    } else {
        // bold size 12 dark blue
        addFont(cell, 'bold', true);
        addFont(cell, 'size', 12);
        addFont(cell, 'color', { rgb: colors.darkBlue });
    }
});

// Cells B64 - I64
const cellsSixtyFour = ['B64', 'C64', 'D64', 'E64', 'F64', 'G64', 'H64', 'I64'];

cellsSixtyFour.forEach((cell, idx) => {
    if (idx % 2 == 0 && idx < 6) {
        // border left thin center bold
        addBorder(cell, 'left', 'thin');
        addAlignment(cell, 'horizontal', 'center');
        addFont(cell, 'bold', true);
    } else if (idx % 2 != 0 && idx < 6) {
        // border right thin
        addBorder(cell, 'right', 'thin');
    } else {
        // border left right thin center bold
        addBorder(cell, 'left', 'thin');
        addBorder(cell, 'right', 'thin');
        addAlignment(cell, 'horizontal', 'center');
        addFont(cell, 'bold', true);
    }
    // border top bottom thin
    addBorder(cell, 'top', 'thin');
    addBorder(cell, 'bottom', 'thin');
});

// Cells A65 - I65
const cellsSixtyFive = ['A65', 'B65', 'C65', 'D65', 'E65', 'F65', 'G65', 'H65', 'I65'];

cellsSixtyFive.forEach((cell, idx) => {
    if (idx == 0) {
        // bold
        addFont(cell, 'bold', true);
    } else if (idx == 1 || idx == 3) {
        // border left bg color cactus
        addBorder(cell, 'left', 'thin');
        addBgColor(cell, colors.cactus);
    } else if (idx == 2 || idx == 4) {
        // border right
        addBorder(cell, 'right', 'thin');
    } else if (idx > 6) {
        // border left right thin bg color cactus
        addBorder(cell, 'left', 'thin');
        addBorder(cell, 'right', 'thin');
        addBgColor(cell, colors.cactus);
    }

    if (idx > 0 && idx != 5 && idx != 6) {
        // border top bottom thin
        addBorder(cell, 'top', 'thin');
        addBorder(cell, 'bottom', 'thin');
    }
});

// Cell L65 center bold
addAlignment('L65', 'horizontal', 'center');
addFont('L65', 'bold', true);

// Cell A66 bold sz 8 right align
addFont('A66', 'bold', true);
addFont('A66', 'sz', 8);
addAlignment('A66', 'horizontal', 'right');

// Cell B66 left bottom top border cactus
addBorder('B66', 'left', 'thin');
addBorder('B66', 'bottom', 'thin');
addBorder('B66', 'top', 'thin');
addBgColor('B66', colors.cactus);

// Cell C66 right bottom top border
addBorder('C66', 'right', 'thin');
addBorder('C66', 'bottom', 'thin');
addBorder('C66', 'top', 'thin');

// Cell L66 - N66
const cellsSixtySix = ['L66', 'M66', 'N66'];

cellsSixtySix.forEach((cell, idx) => {
    if (idx == 0) {
        // border left thin center text bg color peach
        addBorder(cell, 'left', 'thin');
        addAlignment(cell, 'horizontal', 'center');
        addBgColor(cell, colors.peach);
    } else if (idx == 2) {
        // border right
        addBorder(cell, 'right', 'thin');
    }

    // border top bottom thin
    addBorder(cell, 'top', 'thin');
    addBorder(cell, 'bottom', 'thin');
});

// Cells A66 - Q66
const cellsSixtySeven = ['A67', 'B67', 'C67', 'D67', 'E67', 'F67', 'G67', 'H67', 'I67', 'J67', 'K67', 'L67', 'M67', 'N67', 'O67', 'P67', 'Q67'];

cellsSixtySeven.forEach((cell, idx) => {
    // Border top
    addBorder(cell, 'top');
});

// Cells A67 - A71
const cellsSixtyEight = ['A67', 'A68', 'A69', 'A70', 'A71'];

cellsSixtyEight.forEach((cell, idx) => {
    // BOLD
    addFont(cell, 'bold', true);
});

// Cells B68 - B74
const cellsSixtyNine = ['B68', 'B69', 'B70', 'B71', 'B72', 'B73', 'B74'];

cellsSixtyNine.forEach((cell, idx) => {
    // Border left top bottom thin bg cactus
    addBorder(cell, 'left', 'thin');
    addBorder(cell, 'top', 'thin');
    addBorder(cell, 'bottom', 'thin');
    addBgColor(cell, colors.cactus);
});

// Cells C68 - C74 D68 - D74
const cellsSeventy = ['C68', 'C69', 'C70', 'C71', 'C72', 'C73', 'C74', 'D68', 'D69', 'D70', 'D71', 'D72', 'D73', 'D74'];

cellsSeventy.forEach((cell, idx) => {
    // border top bottom
    addBorder(cell, 'top', 'thin');
    addBorder(cell, 'bottom', 'thin');
});

// Cells E68 - E74
const cellsSeventyOne = ['E68', 'E69', 'E70', 'E71', 'E72', 'E73', 'E74'];

cellsSeventyOne.forEach((cell, idx) => {
    // top bottom right border thin
    addBorder(cell, 'top', 'thin');
    addBorder(cell, 'bottom', 'thin');
    addBorder(cell, 'right', 'thin');
});

// Cells F69 - F71
const cellsSeventyTwo = ['F69', 'F70', 'F71'];

cellsSeventyTwo.forEach((cell, idx) => {
    // bold right align
    addFont(cell, 'bold', true);
    addAlignment(cell, 'horizontal', 'right');
});

// Cells H69 - H74 I69 - I74
const cellsSeventyThree = ['H69', 'H70', 'H71', 'H72', 'H73', 'H74', 'I69', 'I70', 'I71', 'I72', 'I73', 'I74'];

cellsSeventyThree.forEach((cell, idx) => {
    if (idx < 6) {
        // border left top bottom thin cactus bg
        addBorder(cell, 'left', 'thin');
        addBorder(cell, 'top', 'thin');
        addBorder(cell, 'bottom', 'thin');
        addBgColor(cell, colors.cactus);
    } else {
        // border top bottom right thin
        addBorder(cell, 'top', 'thin');
        addBorder(cell, 'bottom', 'thin');
        addBorder(cell, 'right', 'thin');
    }
});

// Cell j 70 color dark blue bold 
addFont('J70', 'bold', true);
addFont("J70", "color", { rgb: colors.darkBlue });

// Cell J71 right align bold
addAlignment('J71', 'horizontal', 'right');
addFont('J71', 'bold', true);

// Cells L71 - L74 M71 - M74
const cellsSeventyRow = ['L71', 'L72', 'L73', 'L74', 'M71', 'M72', 'M73', 'M74'];

cellsSeventyRow.forEach((cell, idx) => {
    if (idx < 4) {
        // border left top bottom thin cactus bg
        addBorder(cell, 'left', 'thin');
        addBorder(cell, 'top', 'thin');
        addBorder(cell, 'bottom', 'thin');
        addBgColor(cell, colors.cactus);
    } else {
        // border top bottom right thin
        addBorder(cell, 'top', 'thin');
        addBorder(cell, 'bottom', 'thin');
        addBorder(cell, 'right', 'thin');
    }
});

// Cells A74 - Q74
const cellsSeventyFour = ['A74', 'B74', 'C74', 'D74', 'E74', 'F74', 'G74', 'H74', 'I74', 'J74', 'K74', 'L74', 'M74', 'N74', 'O74', 'P74', 'Q74'];

cellsSeventyFour.forEach((cell, idx) => {
    // border bottom
    addBorder(cell, 'bottom');
});

// Cells A75 - Q75
const cellsSeventyFive = ['A75', 'B75', 'C75', 'D75', 'E75', 'F75', 'G75', 'H75', 'I75', 'J75', 'K75', 'L75', 'M75', 'N75', 'O75', 'P75', 'Q75'];

cellsSeventyFive.forEach((cell, idx) => {
    if (idx == 0) {
        // bold left align
        addFont(cell, 'bold', true);
        addAlignment(cell, 'horizontal', 'left');
    } else {
        // bg color cactus border bottom
        addBgColor(cell, colors.cactus);
        addBorder(cell, 'bottom');
    }
});

// Cells A76 - Q76
const cellsSeventySix = ['A76', 'B76', 'C76', 'D76', 'E76', 'F76', 'G76', 'H76', 'I76', 'J76', 'K76', 'L76', 'M76', 'N76', 'O76', 'P76', 'Q76'];

cellsSeventySix.forEach((cell, idx) => {
    if (idx == 0) {
        // border bottom
        addBorder(cell, 'bottom');
    } else {
        // border bottom thin
        addBorder(cell, 'bottom', 'thin');
    }
});

// Cells B78 - O78
const cellsSeventyEight = ['B78', 'C78', 'D78', 'E78', 'F78', 'G78', 'H78', 'I78', 'J78', 'K78', 'L78', 'M78', 'N78', 'O78'];

cellsSeventyEight.forEach((cell, idx) => {
    // border top thin
    addBorder(cell, 'top', 'thin');
});

// Cells B78 - B89
const cellsSeventyEightRowLeft = ['B78', 'B79', 'B80', 'B81', 'B82', 'B83', 'B84', 'B85', 'B86', 'B87', 'B88', 'B89'];

cellsSeventyEightRowLeft.forEach((cell, idx) => {
    // border left thin
    addBorder(cell, 'left', 'thin');
});

// Cells O78 - O89
const cellsSeventyEightRowRight = ['O78', 'O79', 'O80', 'O81', 'O82', 'O83', 'O84', 'O85', 'O86', 'O87', 'O88', 'O89'];

cellsSeventyEightRowRight.forEach((cell, idx) => {
    // border right thin
    addBorder(cell, 'right', 'thin');
});

// Cells B89 - O89
const cellsEightyNine = ['B89', 'C89', 'D89', 'E89', 'F89', 'G89', 'H89', 'I89', 'J89', 'K89', 'L89', 'M89', 'N89', 'O89'];

cellsEightyNine.forEach((cell, idx) => {
    // border bottom thin
    addBorder(cell, 'bottom', 'thin');
});

// Loop through rows E - N and columns 81 - 87
for (let i = 4; i < 13; i++) {
    for (let j = 80; j < 87; j++) {
        const cell = `${String.fromCharCode(65 + i)}${j}`;
        // border bottom thin
        addBorder(cell, 'bottom', 'thin');
    }
}

// Center underline bold cell B79
addAlignment('B79', 'horizontal', 'center');
addFont('B79', 'underline', true);
addFont('B79', 'bold', true);
// set value to "Warranty Verification Card"
workbook.Sheets[sheetName]["B79"].v = "Warranty Verification Card";

addBgColor("A93", colors.darkBlue);
addFont("A93", "color", { rgb: colors.white });

// Cells Q5 - Q76
const cellsFiveToSeventySix = ['Q5', 'Q6', 'Q7', 'Q8', 'Q9', 'Q10', 'Q11', 'Q12', 'Q13', 'Q14', 'Q15', 'Q16', 'Q17', 'Q18', 'Q19', 'Q20', 'Q21', 'Q22', 'Q23', 'Q24', 'Q25', 'Q26', 'Q27', 'Q28', 'Q29', 'Q30', 'Q31', 'Q32', 'Q33', 'Q34', 'Q35', 'Q36', 'Q37', 'Q38', 'Q39', 'Q40', 'Q41', 'Q42', 'Q43', 'Q44', 'Q45', 'Q46', 'Q47', 'Q48', 'Q49', 'Q50', 'Q51', 'Q52', 'Q53', 'Q54', 'Q55', 'Q56', 'Q57', 'Q58', 'Q59', 'Q60', 'Q61', 'Q62', 'Q63', 'Q64', 'Q65', 'Q66', 'Q67', 'Q68', 'Q69', 'Q70', 'Q71', 'Q72', 'Q73', 'Q74', 'Q75', 'Q76'];

cellsFiveToSeventySix.forEach((cell, idx) => {
    // border right
    addBorder(cell, 'right');
});

// Cells A5 - A76
const cellsA5ToA76 = ['A5', 'A6', 'A7', 'A8', 'A9', 'A10', 'A11', 'A12', 'A13', 'A14', 'A15', 'A16', 'A17', 'A18', 'A19', 'A20', 'A21', 'A22', 'A23', 'A24', 'A25', 'A26', 'A27', 'A28', 'A29', 'A30', 'A31', 'A32', 'A33', 'A34', 'A35', 'A36', 'A37', 'A38', 'A39', 'A40', 'A41', 'A42', 'A43', 'A44', 'A45', 'A46', 'A47', 'A48', 'A49', 'A50', 'A51', 'A52', 'A53', 'A54', 'A55', 'A56', 'A57', 'A58', 'A59', 'A60', 'A61', 'A62', 'A63', 'A64', 'A65', 'A66', 'A67', 'A68', 'A69', 'A70', 'A71', 'A72', 'A73', 'A74', 'A75', 'A76'];

cellsA5ToA76.forEach((cell, idx) => {
    // border left thin
    addBorder(cell, 'left', 'thin');
});


const respectiveCells = [
    // TBC Details
    [
        // formData.tbc.dealerName,
        "C5",
        // formData.tbc.customerName,
        "L5",
        // formData.tbc.boatLength,
        "F6",
        // formData.tbc.boatWidth,
        "L6",
        // formData.tbc.topColor,
        "C7",
        // formData.tbc.length,
        "K7",
        // formData.tbc.sideHeight,
        "P7",
        // formData.tbc.sideColor,
        "C8",
        // formData.tbc.width,
        "K8",
        // formData.tbc.ridge,
        "P8",
        // formData.tbc.frameMaterial,
        "C9",
        // formData.tbc.frameType,
        "K9",
        // formData.tbc.dockType
        "P9"
    ],
    [
        // formData.accessZipper.rearL.x,
        "B12",
        // formData.accessZipper.rearL.y,
        "C12",
        // formData.accessZipper.rearR.x,
        "D12",
        // formData.accessZipper.rearR.y,
        "E12",
        // formData.accessZipper.secondL.x,
        "F12",
        // formData.accessZipper.secondL.y,
        "G12",
        // formData.accessZipper.secondR.x,
        "H12",
        // formData.accessZipper.secondR.y,
        "I12",
        // formData.accessZipper.thirdL.x, 
        "J12",
        // formData.accessZipper.thirdL.y, 
        "K12",
        // formData.accessZipper.thirdR.x,
        "L12",
        // formData.accessZipper.thirdR.y, 
        "M12",
        // formData.accessZipper.frontL.x, 
        "N12",
        // formData.accessZipper.frontL.y, 
        "O12",
        // formData.accessZipper.frontR.x,
        "P12",
        // formData.accessZipper.frontR.y,
        "Q12",
        // formData.accessZipper.drivePipe,
        "B13",
        // formData.accessZipper.liftType 
        "B14"
    ],
    [
        // formData.dholes.customAngle,
        "B27",
        // formData.dholes.regular.rearL.x,
        "B18",
        // formData.dholes.regular.rearL.y,
        "C18",
        // formData.dholes.regular.rearL.l,
        "B20",
        // formData.dholes.regular.rearL.w,
        "C20",
        // formData.dholes.regular.rearR.x,
        "D18",
        // formData.dholes.regular.rearR.y,
        "E18",
        // formData.dholes.regular.rearR.l,
        "D20",
        // formData.dholes.regular.rearR.w,
        "E20",
        // formData.dholes.regular.secondL.x,
        "F18",
        // formData.dholes.regular.secondL.y,
        "G18",
        // formData.dholes.regular.secondL.l,
        "F20",
        // formData.dholes.regular.secondL.w,
        "G20",
        // formData.dholes.regular.secondR.x,
        "H18",
        // formData.dholes.regular.secondR.y,
        "I18",
        // formData.dholes.regular.secondR.l,
        "H20",
        // formData.dholes.regular.secondR.w,
        "I20",
        // formData.dholes.regular.thirdL.x,
        "J18",
        // formData.dholes.regular.thirdL.y,
        "K18",
        // formData.dholes.regular.thirdL.l,
        "J20",
        // formData.dholes.regular.thirdL.w,
        "K20",
        // formData.dholes.regular.thirdR.x,
        "L18",
        // formData.dholes.regular.thirdR.y,
        "M18",
        // formData.dholes.regular.thirdR.l,
        "L20",
        // formData.dholes.regular.thirdR.w,
        "M20",
        // formData.dholes.regular.frontL.x,
        "N18",
        // formData.dholes.regular.frontL.y,
        "O18",
        // formData.dholes.regular.frontL.l,
        "N20",
        // formData.dholes.regular.frontL.w,
        "O20",
        // formData.dholes.regular.frontR.x,
        "P18",
        // formData.dholes.regular.frontR.y,
        "Q18",
        // formData.dholes.regular.frontR.l,
        "P20",
        // formData.dholes.regular.frontR.w,
        "Q20",
        // formData.dholes.open.pile1L.x,
        "B23",
        // formData.dholes.open.pile1L.y,
        "C23",
        // formData.dholes.open.pile1R.x,
        "D23",
        // formData.dholes.open.pile1R.y,
        "E23",
        // formData.dholes.open.pile2L.x,
        "F23",
        // formData.dholes.open.pile2L.y,
        "G23",
        // formData.dholes.open.pile2R.x,
        "H23",
        // formData.dholes.open.pile2R.y,
        "I23",
        // formData.dholes.open.pile3L.x,
        "J23",
        // formData.dholes.open.pile3L.y,
        "K23",
        // formData.dholes.open.pile3R.x,
        "L23",
        // formData.dholes.open.pile3R.y,
        "M23",
        // formData.dholes.open.pile4L.x,
        "N23",
        // formData.dholes.open.pile4L.y,
        "O23",
        // formData.dholes.open.pile4R.x,
        "P23",
        // formData.dholes.open.pile4R.y,
        "Q23",
        // formData.dholes.open.pile5L.x,
        "B26",
        // formData.dholes.open.pile5L.y,
        "C26",
        // formData.dholes.open.pile5R.x,
        "D26",
        // formData.dholes.open.pile5R.y,
        "E26",
        // formData.dholes.open.pile6L.x,
        "F26",
        // formData.dholes.open.pile6L.y,
        "G26",
        // formData.dholes.open.pile6R.x,
        "H26",
        // formData.dholes.open.pile6R.y
        "I26"
    ],
     // Support Cables
     [
        // formData.supportCables.front.x,
        "B30",
        // formData.supportCables.front.y,
        "C30",
        // formData.supportCables.left1.x,
        "D30",
        // formData.supportCables.left1.y,
        "E30",
        // formData.supportCables.left2.x,
        "F30",
        // formData.supportCables.left2.y,
        "G30",
        // formData.supportCables.left3.x,
        "H30",
        // formData.supportCables.left3.y,
        "I30",
        // formData.supportCables.left4.x,
        "J30",
        // formData.supportCables.left4.y,
        "K30",
        // formData.supportCables.left5.x,
        "L30",
        // formData.supportCables.left5.y,
        "M30",
        // formData.supportCables.left6.x,
        "N30",
        // formData.supportCables.left6.y,
        "O30",
        // formData.supportCables.rear.x,
        "B33",
        // formData.supportCables.rear.y,
        "C33",
        // formData.supportCables.right1.x,
        "D33",
        // formData.supportCables.right1.y,
        "E33",
        // formData.supportCables.right2.x,
        "F33",
        // formData.supportCables.right2.y,
        "G33",
        // formData.supportCables.right3.x,
        "H33",
        // formData.supportCables.right3.y,
        "I33",
        // formData.supportCables.right4.x,
        "J33",
        // formData.supportCables.right4.y,
        "K33",
        // formData.supportCables.right5.x,
        "L33",
        // formData.supportCables.right5.y,
        "M33",
        // formData.supportCables.right6.x,
        "N33",
        // formData.supportCables.right6.y,
        "O33",
        // formData.supportCables.hardware,
        "B34",
        // formData.supportCables.notes
        "G34"
    ],
    // Motor
    [
        // formData.motor.power,
        "B35",
        // formData.motor.position
        "G35"
    ],
    // Pilings
    [
        // formData.pilings.rows,
        "B38",
        // formData.pilings.shape,
        "B39",
        // formData.pilings.material,
        "B40",
        // formData.pilings.measurements.left1.x,
        "F39",
        // formData.pilings.measurements.left1.y,
        "G39",
        // formData.pilings.measurements.left1.w,
        "F40",
        // formData.pilings.measurements.left2.x,
        "H39",
        // formData.pilings.measurements.left2.y,
        "I39",
        // formData.pilings.measurements.left2.w,
        "H40",
        // formData.pilings.measurements.left3.x,
        "J39",
        // formData.pilings.measurements.left3.y,
        "K39",
        // formData.pilings.measurements.left3.w,
        "J40",
        // formData.pilings.measurements.left4.x,
        "L39",
        // formData.pilings.measurements.left4.y,
        "M39",
        // formData.pilings.measurements.left4.w,
        "L40",
        // formData.pilings.measurements.left5.x,
        "N39",
        // formData.pilings.measurements.left5.y,
        "O39",
        // formData.pilings.measurements.left5.w,
        "N40",
        // formData.pilings.measurements.left6.x,
        "P39",
        // formData.pilings.measurements.left6.y,
        "Q39",
        // formData.pilings.measurements.left6.w,
        "P40",
        // formData.pilings.measurements.right1.x,
        "F43",
        // formData.pilings.measurements.right1.y,
        "G43",
        // formData.pilings.measurements.right1.w,
        "F44",
        // formData.pilings.measurements.right2.x,
        "H43",
        // formData.pilings.measurements.right2.y,
        "I43",
        // formData.pilings.measurements.right2.w,
        "H44",
        // formData.pilings.measurements.right3.x,
        "J43",
        // formData.pilings.measurements.right3.y,
        "K43",
        // formData.pilings.measurements.right3.w,
        "J44",
        // formData.pilings.measurements.right4.x,
        "L43",
        // formData.pilings.measurements.right4.y,
        "M43",
        // formData.pilings.measurements.right4.w,
        "L44",
        // formData.pilings.measurements.right5.x,
        "N43",
        // formData.pilings.measurements.right5.y,
        "O43",
        // formData.pilings.measurements.right5.w,
        "N44",
        // formData.pilings.measurements.right6.x,
        "P43",
        // formData.pilings.measurements.right6.y,
        "Q43",
        // formData.pilings.measurements.right6.w,
        "P44",
        // formData.pilings.measurements.itoi1,
        "F45",
        // formData.pilings.measurements.itoi2,
        "H45",
        // formData.pilings.measurements.itoi3,
        "J45",
        // formData.pilings.measurements.itoi4,
        "L45",
        // formData.pilings.measurements.itoi5,
        "N45",
        // formData.pilings.measurements.itoi6,
        "P45",
        // formData.pilings.cover.height,
        "D46",
        // formData.pilings.cover.width,
        "G46",
        // formData.pilings.cover.quantity,
        "J46",
        // formData.pilings.cover.type,
        "L46",
        // formData.pilings.cover.color
        "P46"
    ],
    // Brackets
    [
        // formData.brackets.reaches[0].ty,
        "B49",
        // formData.brackets.reaches[0].values[0],
        "F49",
        // formData.brackets.reaches[0].values[1],
        "H49",
        // formData.brackets.reaches[0].values[2],
        "J49",
        // formData.brackets.reaches[0].values[3],
        "L49",
        // formData.brackets.reaches[0].values[4],
        "N49",
        // formData.brackets.reaches[0].values[5],
        "P49",
        // formData.brackets.reaches[1].ty,
        "B50",
        // formData.brackets.reaches[1].values[0],
        "F50",
        // formData.brackets.reaches[1].values[1],
        "H50",
        // formData.brackets.reaches[1].values[2],
        "J50",
        // formData.brackets.reaches[1].values[3],
        "L50",
        // formData.brackets.reaches[1].values[4],
        "N50",
        // formData.brackets.reaches[1].values[5],
        "P50",
        // formData.brackets.reaches[2].ty,
        "B51",
        // formData.brackets.reaches[2].values[0],
        "F51",
        // formData.brackets.reaches[2].values[1],
        "H51",
        // formData.brackets.reaches[2].values[2],
        "J51",
        // formData.brackets.reaches[2].values[3],
        "L51",
        // formData.brackets.reaches[2].values[4],
        "N51",
        // formData.brackets.reaches[2].values[5],
        "P51",
        // formData.brackets.heights[0].ty,
        "B53",
        // formData.brackets.heights[0].values[0],
        "F53",
        // formData.brackets.heights[0].values[1],
        "H53",
        // formData.brackets.heights[0].values[2],
        "J53",
        // formData.brackets.heights[0].values[3],
        "L53",
        // formData.brackets.heights[0].values[4],
        "N53",
        // formData.brackets.heights[0].values[5],
        "P53",
        // formData.brackets.heights[1].ty,
        "B54",
        // formData.brackets.heights[1].values[0],
        "F54",
        // formData.brackets.heights[1].values[1],
        "H54",
        // formData.brackets.heights[1].values[2],
        "J54",
        // formData.brackets.heights[1].values[3],
        "L54",
        // formData.brackets.heights[1].values[4],
        "N54",
        // formData.brackets.heights[1].values[5],
        "P54",
        // formData.brackets.hardware,
        "B55",
        // formData.brackets.attachment,
        "B56",
        // formData.brackets.notes
        "G55"
    ],
    // Custom
    [
        // formData.custom.boot.l,
        "B58",
        // formData.custom.boot.w,
        "D58",
        // formData.custom.boot.h,
        "F58",
        // formData.custom.boot.x,
        "J58",
        // formData.custom.boot.y,
        "K58",
        // formData.custom.boot.color,
        "L58",
        // formData.custom.nose.bend,
        "B60",
        // formData.custom.nose.width,
        "D60",
        // formData.custom.rear.bend,
        "B61",
        // formData.custom.rear.width,
        "D61",
        // formData.custom.ridge.enabled,
        "B62",
        // formData.custom.frame.enabled,
        "B63",
        // formData.custom.openPassThru.l,
        "B65",
        // formData.custom.openPassThru.w,
        "D65",
        // formData.custom.openPassThru.h, ????
        // formData.custom.openPassThru.x,
        "H65",
        // formData.custom.openPassThru.y,
        "I65",
        // formData.custom.sideRtdCable,
        "B66",
        // formData.custom.consoleWidth
        "L66"
    ],
    // Additional
    [
        // formData.additional.midframe,
        "B68",
        // formData.additional.remotes,
        "B69",
        // formData.additional.ridgeSupports,
        "B70",
        // formData.additional.parts[0].name,
        "B71",
        // formData.additional.parts[0].amount,
        "H71",
        // formData.additional.parts[0].cost,
        "L71",
        // formData.additional.parts[1].name,
        "B72",
        // formData.additional.parts[1].amount,
        "H72",
        // formData.additional.parts[1].cost,
        "L72",
        // formData.additional.parts[2].name,
        "B73",
        // formData.additional.parts[2].amount,
        "H73",
        // formData.additional.parts[2].cost,
        "L73",
        // formData.additional.parts[3].name,
        "B74",
        // formData.additional.parts[3].amount,
        "H74",
        // formData.additional.parts[3].cost,
        "L74",
        // formData.additional.notes
        "B75"
    ]
];

function getSafe(fn, defaultVal="") {
    try {
      return fn();
    } catch (e) {
      return defaultVal;
    }
}
window.createXLSXFile = function (formData) {
    const outData = [
        // TBC Details
        [
            formData.tbc.dealerName,
            formData.tbc.customerName,
            formData.tbc.boatLength,
            formData.tbc.boatWidth,
            formData.tbc.topColor,
            formData.tbc.length,
            formData.tbc.sideHeight,
            formData.tbc.sideColor,
            formData.tbc.width,
            formData.tbc.ridge,
            formData.tbc.frameMaterial,
            formData.tbc.frameType,
            formData.tbc.dockType
        ],
        // Access Zipper
        [
            formData.accessZipper.rearL.x,
            formData.accessZipper.rearL.y,
            formData.accessZipper.rearR.x,
            formData.accessZipper.rearR.y,
            formData.accessZipper.secondL.x,
            formData.accessZipper.secondL.y,
            formData.accessZipper.secondR.x,
            formData.accessZipper.secondR.y,
            formData.accessZipper.thirdL.x, 
            formData.accessZipper.thirdL.y, 
            formData.accessZipper.thirdR.x, 
            formData.accessZipper.thirdR.y, 
            formData.accessZipper.frontL.x, 
            formData.accessZipper.frontL.y, 
            formData.accessZipper.frontR.x, 
            formData.accessZipper.frontR.y, 
            formData.accessZipper.drivePipe,
            formData.accessZipper.liftType 
        ],
        // Dholes
        [
            formData.dholes.customAngle,
            formData.dholes.regular.rearL.x,
            formData.dholes.regular.rearL.y,
            formData.dholes.regular.rearL.l,
            formData.dholes.regular.rearL.w,
            formData.dholes.regular.rearR.x,
            formData.dholes.regular.rearR.y,
            formData.dholes.regular.rearR.l,
            formData.dholes.regular.rearR.w,
            formData.dholes.regular.secondL.x,
            formData.dholes.regular.secondL.y,
            formData.dholes.regular.secondL.l,
            formData.dholes.regular.secondL.w,
            formData.dholes.regular.secondR.x,
            formData.dholes.regular.secondR.y,
            formData.dholes.regular.secondR.l,
            formData.dholes.regular.secondR.w,
            formData.dholes.regular.thirdL.x,
            formData.dholes.regular.thirdL.y,
            formData.dholes.regular.thirdL.l,
            formData.dholes.regular.thirdL.w,
            formData.dholes.regular.thirdR.x,
            formData.dholes.regular.thirdR.y,
            formData.dholes.regular.thirdR.l,
            formData.dholes.regular.thirdR.w,
            formData.dholes.regular.frontL.x,
            formData.dholes.regular.frontL.y,
            formData.dholes.regular.frontL.l,
            formData.dholes.regular.frontL.w,
            formData.dholes.regular.frontR.x,
            formData.dholes.regular.frontR.y,
            formData.dholes.regular.frontR.l,
            formData.dholes.regular.frontR.w,
            formData.dholes.open.pile1L.x,
            formData.dholes.open.pile1L.y,
            formData.dholes.open.pile1R.x,
            formData.dholes.open.pile1R.y,
            formData.dholes.open.pile2L.x,
            formData.dholes.open.pile2L.y,
            formData.dholes.open.pile2R.x,
            formData.dholes.open.pile2R.y,
            formData.dholes.open.pile3L.x,
            formData.dholes.open.pile3L.y,
            formData.dholes.open.pile3R.x,
            formData.dholes.open.pile3R.y,
            formData.dholes.open.pile4L.x,
            formData.dholes.open.pile4L.y,
            formData.dholes.open.pile4R.x,
            formData.dholes.open.pile4R.y,
            formData.dholes.open.pile5L.x,
            formData.dholes.open.pile5L.y,
            formData.dholes.open.pile5R.x,
            formData.dholes.open.pile5R.y,
            formData.dholes.open.pile6L.x,
            formData.dholes.open.pile6L.y,
            formData.dholes.open.pile6R.x,
            formData.dholes.open.pile6R.y
        ],
        // Support Cables
        [
            formData.supportCables.front.x,
            formData.supportCables.front.y,
            formData.supportCables.left1.x,
            formData.supportCables.left1.y,
            formData.supportCables.left2.x,
            formData.supportCables.left2.y,
            formData.supportCables.left3.x,
            formData.supportCables.left3.y,
            formData.supportCables.left4.x,
            formData.supportCables.left4.y,
            formData.supportCables.left5.x,
            formData.supportCables.left5.y,
            formData.supportCables.left6.x,
            formData.supportCables.left6.y,
            formData.supportCables.rear.x,
            formData.supportCables.rear.y,
            formData.supportCables.right1.x,
            formData.supportCables.right1.y,
            formData.supportCables.right2.x,
            formData.supportCables.right2.y,
            formData.supportCables.right3.x,
            formData.supportCables.right3.y,
            formData.supportCables.right4.x,
            formData.supportCables.right4.y,
            formData.supportCables.right5.x,
            formData.supportCables.right5.y,
            formData.supportCables.right6.x,
            formData.supportCables.right6.y,
            formData.supportCables.right6.y,
            formData.supportCables.hardware,
            formData.supportCables.notes
        ],
        // Motor
        [
            formData.motor.power,
            formData.motor.position
        ],
        // Pilings
        [
            formData.pilings.rows,
            formData.pilings.shape,
            formData.pilings.material,
            formData.pilings.measurements.left1.x,
            formData.pilings.measurements.left1.y,
            formData.pilings.measurements.left1.w,
            formData.pilings.measurements.left2.x,
            formData.pilings.measurements.left2.y,
            formData.pilings.measurements.left2.w,
            formData.pilings.measurements.left3.x,
            formData.pilings.measurements.left3.y,
            formData.pilings.measurements.left3.w,
            formData.pilings.measurements.left4.x,
            formData.pilings.measurements.left4.y,
            formData.pilings.measurements.left4.w,
            formData.pilings.measurements.left5.x,
            formData.pilings.measurements.left5.y,
            formData.pilings.measurements.left5.w,
            formData.pilings.measurements.left6.x,
            formData.pilings.measurements.left6.y,
            formData.pilings.measurements.left6.w,
            formData.pilings.measurements.right1.x,
            formData.pilings.measurements.right1.y,
            formData.pilings.measurements.right1.w,
            formData.pilings.measurements.right2.x,
            formData.pilings.measurements.right2.y,
            formData.pilings.measurements.right2.w,
            formData.pilings.measurements.right3.x,
            formData.pilings.measurements.right3.y,
            formData.pilings.measurements.right3.w,
            formData.pilings.measurements.right4.x,
            formData.pilings.measurements.right4.y,
            formData.pilings.measurements.right4.w,
            formData.pilings.measurements.right5.x,
            formData.pilings.measurements.right5.y,
            formData.pilings.measurements.right5.w,
            formData.pilings.measurements.right6.x,
            formData.pilings.measurements.right6.y,
            formData.pilings.measurements.right6.w,
            formData.pilings.measurements.itoi1,
            formData.pilings.measurements.itoi2,
            formData.pilings.measurements.itoi3,
            formData.pilings.measurements.itoi4,
            formData.pilings.measurements.itoi5,
            formData.pilings.measurements.itoi6,
            formData.pilings.cover.height,
            formData.pilings.cover.width,
            formData.pilings.cover.quantity,
            formData.pilings.cover.type,
            formData.pilings.cover.color
        ],
        // Brackets
        [
            formData.brackets.reaches[0].ty,
            formData.brackets.reaches[0].values[0],
            formData.brackets.reaches[0].values[1],
            formData.brackets.reaches[0].values[2],
            formData.brackets.reaches[0].values[3],
            formData.brackets.reaches[0].values[4],
            formData.brackets.reaches[0].values[5],
            formData.brackets.reaches[1].ty,
            formData.brackets.reaches[1].values[0],
            formData.brackets.reaches[1].values[1],
            formData.brackets.reaches[1].values[2],
            formData.brackets.reaches[1].values[3],
            formData.brackets.reaches[1].values[4],
            formData.brackets.reaches[1].values[5],
            formData.brackets.reaches[2].ty,
            formData.brackets.reaches[2].values[0],
            formData.brackets.reaches[2].values[1],
            formData.brackets.reaches[2].values[2],
            formData.brackets.reaches[2].values[3],
            formData.brackets.reaches[2].values[4],
            formData.brackets.reaches[2].values[5],
            formData.brackets.heights[0].ty,
            formData.brackets.heights[0].values[0],
            formData.brackets.heights[0].values[1],
            formData.brackets.heights[0].values[2],
            formData.brackets.heights[0].values[3],
            formData.brackets.heights[0].values[4],
            formData.brackets.heights[0].values[5],
            formData.brackets.heights[1].ty,
            formData.brackets.heights[1].values[0],
            formData.brackets.heights[1].values[1],
            formData.brackets.heights[1].values[2],
            formData.brackets.heights[1].values[3],
            formData.brackets.heights[1].values[4],
            formData.brackets.heights[1].values[5],
            formData.brackets.hardware,
            formData.brackets.attachment,
            formData.brackets.notes
        ],
        // Custom
        [
            formData.custom.boot.l,
            formData.custom.boot.w,
            formData.custom.boot.h,
            formData.custom.boot.x,
            formData.custom.boot.y,
            formData.custom.boot.color,
            formData.custom.nose.bend,
            formData.custom.nose.width,
            formData.custom.rear.bend,
            formData.custom.rear.width,
            formData.custom.ridge.enabled,
            formData.custom.frame.enabled,
            formData.custom.openPassThru.l,
            formData.custom.openPassThru.w,
            formData.custom.openPassThru.x,
            formData.custom.openPassThru.y,
            formData.custom.sideRtdCable,
            formData.custom.consoleWidth
        ],
        // Additional
        [
            formData.additional.midframe,
            formData.additional.remotes,
            formData.additional.ridgeSupports,
            formData.additional.parts[0].name,
            formData.additional.parts[0].amount,
            formData.additional.parts[0].cost,
            formData.additional.parts[1].name,
            formData.additional.parts[1].amount,
            formData.additional.parts[1].cost,
            formData.additional.parts[2].name,
            formData.additional.parts[2].amount,
            formData.additional.parts[2].cost,
            formData.additional.parts[3].name,
            formData.additional.parts[3].amount,
            formData.additional.parts[3].cost,
            formData.additional.notes
        ]
    ];

    // Loop through respectiveCells and set the values
    for (let i = 0; i < respectiveCells.length; i++) {
        for (let j = 0; j < respectiveCells[i].length; j++) {
            let cell = respectiveCells[i][j];
            let value = outData[i][j];
            if (value) {
                workbook.Sheets[sheetName][cell].v = value;
            }
        }
    }


    XLSX.writeFile(workbook, 'out.xlsx');
}

// Loads the form data from an XLSX file
window.loadXLSXFile = function(file) {
    let formData = {
        tbc: {
            dealerName: "",
            customerName: "",
            boatLength: "",
            boatWidth: "",
            topColor: "",
            length: "",
            sideHeight: "",
            sideColor: "",
            width: "",
            ridge: "",
            frameMaterial: "",
            frameType: "",
            dockType: "",
        },
        accessZipper: {
            rearL: { x: "", y: "" },
            rearR: { x: "", y: "" },
            secondL: { x: "", y: "" },
            secondR: { x: "", y: "" },
            thirdL: { x: "", y: "" },
            thirdR: { x: "", y: "" },
            frontL: { x: "", y: "" },
            frontR: { x: "", y: "" },
            drivePipe: "",
            liftType: ""
        },
        dholes: {
            regular: {
                rearL: { x: "", y: "", l: "", w: "" },
                rearR: { x: "", y: "", l: "", w: "" },
                secondL: { x: "", y: "", l: "", w: "" },
                secondR: { x: "", y: "", l: "", w: "" },
                thirdL: { x: "", y: "", l: "", w: "" },
                thirdR: { x: "", y: "", l: "", w: "" },
                frontL: { x: "", y: "", l: "", w: "" },
                frontR: { x: "", y: "", l: "", w: "" }
            },
            open: {
                pile1L: { x: "", y: "" },
                pile1R: { x: "", y: "" },
                pile2L: { x: "", y: "" },
                pile2R: { x: "", y: "" },
                pile3L: { x: "", y: "" },
                pile3R: { x: "", y: "" },
                pile4L: { x: "", y: "" },
                pile4R: { x: "", y: "" },
                pile5L: { x: "", y: "" },
                pile5R: { x: "", y: "" },
                pile6L: { x: "", y: "" },
                pile6R: { x: "", y: "" }
            },
            customAngle: ""
        },
        supportCables: {
            front: { x: "", y: "" },
            left1: { x: "", y: "" },
            left2: { x: "", y: "" },
            left3: { x: "", y: "" },
            left4: { x: "", y: "" },
            left5: { x: "", y: "" },
            left6: { x: "", y: "" },
            rear: { x: "", y: "" },
            right1: { x: "", y: "" },
            right2: { x: "", y: "" },
            right3: { x: "", y: "" },
            right4: { x: "", y: "" },
            right5: { x: "", y: "" },
            right6: { x: "", y: "" },
            hardware: "",
            notes: ""
        },
        motor: {
            power: "",
            position: ""
        },
        pilings: {
            rows: "",
            shape: "",
            material: "",
            measurements: {
                left1: { x: "", y: "", w: "" },
                left2: { x: "", y: "", w: "" },
                left3: { x: "", y: "", w: "" },
                left4: { x: "", y: "", w: "" },
                left5: { x: "", y: "", w: "" },
                left6: { x: "", y: "", w: "" },
                right1: { x: "", y: "", w: "" },
                right2: { x: "", y: "", w: "" },
                right3: { x: "", y: "", w: "" },
                right4: { x: "", y: "", w: "" },
                right5: { x: "", y: "", w: "" },
                right6: { x: "", y: "", w: "" },
                itoi1: "",
                itoi2: "",
                itoi3: "",
                itoi4: "",
                itoi5: "",
                itoi6: ""
            },
            cover: {
                height: "",
                width: "",
                quantity: "",
                type: "",
                color: ""
            }
        },
        brackets: {
            reaches: [
                { ty: "", values: [ "", "", "", "", "", "" ] },
                { ty: "", values: [ "", "", "", "", "", "" ] },
                { ty: "", values: [ "", "", "", "", "", "" ] }   
            ],
            heights: [
                { ty: "", values: [ "", "", "", "", "", "" ] },
                { ty: "", values: [ "", "", "", "", "", "" ] }
            ],
            hardware: "",
            attachment: "",
            notes: ""
        },
        custom: {
            boot: { l: "", w: "", h: "", x: "", y: "", color: "" },
            nose: { bend: "", width: "" },
            rear: { bend: "", width: "" },
            ridge: {
                enabled: "",
                image: ""
            },
            frame: {
                enabled: "",
                image: ""
            },
            openPassThru: { l: "", w: "", h: "", x: "", y: "" },
            sideRtdCable: "",
            consoleWidth: ""
        },
        additional: {
            midframe: "",
            remotes: "",
            ridgeSupports: "",
            parts: [
                { name: "", amount: "", cost: "" },
                { name: "", amount: "", cost: "" },
                { name: "", amount: "", cost: "" },
                { name: "", amount: "", cost: "" }
            ],
            notes: ""
        }
    };

    // Default form data requested
    if (file === null) {
        return formData;
    }

    // Load from XLSX
    const loadedWorkbook = XLSX.read(file, { type: "array" });
    const loadedSheetName = loadedWorkbook.SheetNames[0];

    // TBC Details
    formData.tbc.dealerName = workbook.Sheets[loadedSheetName][respectiveCells[0][0]].v;
    formData.tbc.customerName  = workbook.Sheets[loadedSheetName][respectiveCells[0][1]].v;
    formData.tbc.boatLength = workbook.Sheets[loadedSheetName][respectiveCells[0][2]].v;
    formData.tbc.boatWidth = workbook.Sheets[loadedSheetName][respectiveCells[0][3]].v;
    formData.tbc.topColor = workbook.Sheets[loadedSheetName][respectiveCells[0][4]].v;
    formData.tbc.length = workbook.Sheets[loadedSheetName][respectiveCells[0][5]].v;
    formData.tbc.sideHeight = workbook.Sheets[loadedSheetName][respectiveCells[0][6]].v;
    formData.tbc.sideColor = workbook.Sheets[loadedSheetName][respectiveCells[0][7]].v;
    formData.tbc.width = workbook.Sheets[loadedSheetName][respectiveCells[0][8]].v;
    formData.tbc.ridge = workbook.Sheets[loadedSheetName][respectiveCells[0][9]].v;
    formData.tbc.frameMaterial = workbook.Sheets[loadedSheetName][respectiveCells[0][10]].v;
    formData.tbc.frameType = workbook.Sheets[loadedSheetName][respectiveCells[0][11]].v;
    formData.tbc.dockType = workbook.Sheets[loadedSheetName][respectiveCells[0][12]].v;

    // Access Zipper
    formData.accessZipper.rearL.x = workbook.Sheets[loadedSheetName][respectiveCells[1][0]].v;
    formData.accessZipper.rearL.y = workbook.Sheets[loadedSheetName][respectiveCells[1][1]].v;
    formData.accessZipper.rearR.x = workbook.Sheets[loadedSheetName][respectiveCells[1][2]].v;
    formData.accessZipper.rearR.y = workbook.Sheets[loadedSheetName][respectiveCells[1][3]].v;
    formData.accessZipper.secondL.x = workbook.Sheets[loadedSheetName][respectiveCells[1][4]].v;
    formData.accessZipper.secondL.y = workbook.Sheets[loadedSheetName][respectiveCells[1][5]].v;
    formData.accessZipper.secondR.x = workbook.Sheets[loadedSheetName][respectiveCells[1][6]].v;
    formData.accessZipper.secondR.y = workbook.Sheets[loadedSheetName][respectiveCells[1][7]].v;
    formData.accessZipper.thirdL.x = workbook.Sheets[loadedSheetName][respectiveCells[1][8]].v;
    formData.accessZipper.thirdL.y = workbook.Sheets[loadedSheetName][respectiveCells[1][9]].v;
    formData.accessZipper.thirdR.x = workbook.Sheets[loadedSheetName][respectiveCells[1][10]].v;
    formData.accessZipper.thirdR.y = workbook.Sheets[loadedSheetName][respectiveCells[1][11]].v;
    formData.accessZipper.frontL.x = workbook.Sheets[loadedSheetName][respectiveCells[1][12]].v;
    formData.accessZipper.frontL.y = workbook.Sheets[loadedSheetName][respectiveCells[1][13]].v;
    formData.accessZipper.frontR.x = workbook.Sheets[loadedSheetName][respectiveCells[1][14]].v;
    formData.accessZipper.frontR.y = workbook.Sheets[loadedSheetName][respectiveCells[1][15]].v;
    formData.accessZipper.drivePipe = workbook.Sheets[loadedSheetName][respectiveCells[1][16]].v;
    formData.accessZipper.liftType = workbook.Sheets[loadedSheetName][respectiveCells[1][17]].v;

    // Dholes
    formData.dholes.customAngle = workbook.Sheets[loadedSheetName][respectiveCells[2][0]].v;
    formData.dholes.regular.rearL.x = workbook.Sheets[loadedSheetName][respectiveCells[2][1]].v;
    formData.dholes.regular.rearL.y = workbook.Sheets[loadedSheetName][respectiveCells[2][2]].v;
    formData.dholes.regular.rearL.l = workbook.Sheets[loadedSheetName][respectiveCells[2][3]].v;
    formData.dholes.regular.rearL.w = workbook.Sheets[loadedSheetName][respectiveCells[2][4]].v;
    formData.dholes.regular.rearR.x = workbook.Sheets[loadedSheetName][respectiveCells[2][5]].v;
    formData.dholes.regular.rearR.y = workbook.Sheets[loadedSheetName][respectiveCells[2][6]].v;
    formData.dholes.regular.rearR.l = workbook.Sheets[loadedSheetName][respectiveCells[2][7]].v;
    formData.dholes.regular.rearR.w = workbook.Sheets[loadedSheetName][respectiveCells[2][8]].v;
    formData.dholes.regular.secondL.x = workbook.Sheets[loadedSheetName][respectiveCells[2][9]].v;
    formData.dholes.regular.secondL.y = workbook.Sheets[loadedSheetName][respectiveCells[2][10]].v;
    formData.dholes.regular.secondL.l = workbook.Sheets[loadedSheetName][respectiveCells[2][11]].v;
    formData.dholes.regular.secondL.w = workbook.Sheets[loadedSheetName][respectiveCells[2][12]].v;
    formData.dholes.regular.secondR.x = workbook.Sheets[loadedSheetName][respectiveCells[2][13]].v;
    formData.dholes.regular.secondR.y = workbook.Sheets[loadedSheetName][respectiveCells[2][14]].v;
    formData.dholes.regular.secondR.l = workbook.Sheets[loadedSheetName][respectiveCells[2][15]].v;
    formData.dholes.regular.secondR.w = workbook.Sheets[loadedSheetName][respectiveCells[2][16]].v;
    formData.dholes.regular.thirdL.x = workbook.Sheets[loadedSheetName][respectiveCells[2][17]].v;
    formData.dholes.regular.thirdL.y = workbook.Sheets[loadedSheetName][respectiveCells[2][18]].v;
    formData.dholes.regular.thirdL.l = workbook.Sheets[loadedSheetName][respectiveCells[2][19]].v;
    formData.dholes.regular.thirdL.w = workbook.Sheets[loadedSheetName][respectiveCells[2][20]].v;
    formData.dholes.regular.thirdR.x = workbook.Sheets[loadedSheetName][respectiveCells[2][21]].v;
    formData.dholes.regular.thirdR.y = workbook.Sheets[loadedSheetName][respectiveCells[2][22]].v;
    formData.dholes.regular.thirdR.l = workbook.Sheets[loadedSheetName][respectiveCells[2][23]].v;
    formData.dholes.regular.thirdR.w = workbook.Sheets[loadedSheetName][respectiveCells[2][24]].v;
    formData.dholes.regular.frontL.x = workbook.Sheets[loadedSheetName][respectiveCells[2][25]].v;
    formData.dholes.regular.frontL.y = workbook.Sheets[loadedSheetName][respectiveCells[2][26]].v;
    formData.dholes.regular.frontL.l = workbook.Sheets[loadedSheetName][respectiveCells[2][27]].v;
    formData.dholes.regular.frontL.w = workbook.Sheets[loadedSheetName][respectiveCells[2][28]].v;
    formData.dholes.regular.frontR.x = workbook.Sheets[loadedSheetName][respectiveCells[2][29]].v;
    formData.dholes.regular.frontR.y = workbook.Sheets[loadedSheetName][respectiveCells[2][30]].v;
    formData.dholes.regular.frontR.l = workbook.Sheets[loadedSheetName][respectiveCells[2][31]].v;
    formData.dholes.regular.frontR.w = workbook.Sheets[loadedSheetName][respectiveCells[2][32]].v;
    formData.dholes.open.pile1L.x = workbook.Sheets[loadedSheetName][respectiveCells[2][33]].v;
    formData.dholes.open.pile1L.y = workbook.Sheets[loadedSheetName][respectiveCells[2][34]].v;
    formData.dholes.open.pile1R.x = workbook.Sheets[loadedSheetName][respectiveCells[2][35]].v;
    formData.dholes.open.pile1R.y = workbook.Sheets[loadedSheetName][respectiveCells[2][36]].v;
    formData.dholes.open.pile2L.x = workbook.Sheets[loadedSheetName][respectiveCells[2][37]].v;
    formData.dholes.open.pile2L.y = workbook.Sheets[loadedSheetName][respectiveCells[2][38]].v;
    formData.dholes.open.pile2R.x = workbook.Sheets[loadedSheetName][respectiveCells[2][39]].v;
    formData.dholes.open.pile2R.y = workbook.Sheets[loadedSheetName][respectiveCells[2][40]].v;
    formData.dholes.open.pile3L.x = workbook.Sheets[loadedSheetName][respectiveCells[2][41]].v;
    formData.dholes.open.pile3L.y = workbook.Sheets[loadedSheetName][respectiveCells[2][42]].v;
    formData.dholes.open.pile3R.x = workbook.Sheets[loadedSheetName][respectiveCells[2][43]].v;
    formData.dholes.open.pile3R.y = workbook.Sheets[loadedSheetName][respectiveCells[2][44]].v;
    formData.dholes.open.pile4L.x = workbook.Sheets[loadedSheetName][respectiveCells[2][45]].v;
    formData.dholes.open.pile4L.y = workbook.Sheets[loadedSheetName][respectiveCells[2][46]].v;
    formData.dholes.open.pile4R.x = workbook.Sheets[loadedSheetName][respectiveCells[2][47]].v;
    formData.dholes.open.pile4R.y = workbook.Sheets[loadedSheetName][respectiveCells[2][48]].v;
    formData.dholes.open.pile5L.x = workbook.Sheets[loadedSheetName][respectiveCells[2][49]].v;
    formData.dholes.open.pile5L.y = workbook.Sheets[loadedSheetName][respectiveCells[2][50]].v;
    formData.dholes.open.pile5R.x = workbook.Sheets[loadedSheetName][respectiveCells[2][51]].v;
    formData.dholes.open.pile5R.y = workbook.Sheets[loadedSheetName][respectiveCells[2][52]].v;
    formData.dholes.open.pile6L.x = workbook.Sheets[loadedSheetName][respectiveCells[2][53]].v;
    formData.dholes.open.pile6L.y = workbook.Sheets[loadedSheetName][respectiveCells[2][54]].v;
    formData.dholes.open.pile6R.x = workbook.Sheets[loadedSheetName][respectiveCells[2][55]].v;
    formData.dholes.open.pile6R.y = workbook.Sheets[loadedSheetName][respectiveCells[2][56]].v;

    // Support Cables
    formData.supportCables.front.x = workbook.Sheets[loadedSheetName][respectiveCells[3][0]].v;
    formData.supportCables.front.y = workbook.Sheets[loadedSheetName][respectiveCells[3][1]].v;
    formData.supportCables.left1.x = workbook.Sheets[loadedSheetName][respectiveCells[3][2]].v;
    formData.supportCables.left1.y = workbook.Sheets[loadedSheetName][respectiveCells[3][3]].v;
    formData.supportCables.left2.x = workbook.Sheets[loadedSheetName][respectiveCells[3][4]].v;
    formData.supportCables.left2.y = workbook.Sheets[loadedSheetName][respectiveCells[3][5]].v;
    formData.supportCables.left3.x = workbook.Sheets[loadedSheetName][respectiveCells[3][6]].v;
    formData.supportCables.left3.y = workbook.Sheets[loadedSheetName][respectiveCells[3][7]].v;
    formData.supportCables.left4.x = workbook.Sheets[loadedSheetName][respectiveCells[3][8]].v;
    formData.supportCables.left4.y = workbook.Sheets[loadedSheetName][respectiveCells[3][9]].v;
    formData.supportCables.left5.x = workbook.Sheets[loadedSheetName][respectiveCells[3][10]].v;
    formData.supportCables.left5.y = workbook.Sheets[loadedSheetName][respectiveCells[3][11]].v;
    formData.supportCables.left6.x = workbook.Sheets[loadedSheetName][respectiveCells[3][12]].v;
    formData.supportCables.left6.y = workbook.Sheets[loadedSheetName][respectiveCells[3][13]].v;
    formData.supportCables.rear.x = workbook.Sheets[loadedSheetName][respectiveCells[3][14]].v;
    formData.supportCables.rear.y = workbook.Sheets[loadedSheetName][respectiveCells[3][15]].v;
    formData.supportCables.right1.x = workbook.Sheets[loadedSheetName][respectiveCells[3][16]].v;
    formData.supportCables.right1.y = workbook.Sheets[loadedSheetName][respectiveCells[3][17]].v;
    formData.supportCables.right2.x = workbook.Sheets[loadedSheetName][respectiveCells[3][18]].v;
    formData.supportCables.right2.y = workbook.Sheets[loadedSheetName][respectiveCells[3][19]].v;
    formData.supportCables.right3.x = workbook.Sheets[loadedSheetName][respectiveCells[3][20]].v;
    formData.supportCables.right3.y = workbook.Sheets[loadedSheetName][respectiveCells[3][21]].v;
    formData.supportCables.right4.x = workbook.Sheets[loadedSheetName][respectiveCells[3][22]].v;
    formData.supportCables.right4.y = workbook.Sheets[loadedSheetName][respectiveCells[3][23]].v;
    formData.supportCables.right5.x = workbook.Sheets[loadedSheetName][respectiveCells[3][24]].v;
    formData.supportCables.right5.y = workbook.Sheets[loadedSheetName][respectiveCells[3][25]].v;
    formData.supportCables.right6.x = workbook.Sheets[loadedSheetName][respectiveCells[3][26]].v;
    formData.supportCables.right6.y = workbook.Sheets[loadedSheetName][respectiveCells[3][27]].v;
    formData.supportCables.hardware = workbook.Sheets[loadedSheetName][respectiveCells[3][28]].v;
    formData.supportCables.notes = workbook.Sheets[loadedSheetName][respectiveCells[3][29]].v;

    // Motor
    formData.motor.power = workbook.Sheets[loadedSheetName][respectiveCells[4][0]].v;
    formData.motor.position = workbook.Sheets[loadedSheetName][respectiveCells[4][1]].v;

    // Pilings
    formData.pilings.rows = workbook.Sheets[loadedSheetName][respectiveCells[5][0]].v;
    formData.pilings.shape = workbook.Sheets[loadedSheetName][respectiveCells[5][1]].v;
    formData.pilings.material = workbook.Sheets[loadedSheetName][respectiveCells[5][2]].v;
    formData.pilings.measurements.left1.x = workbook.Sheets[loadedSheetName][respectiveCells[5][3]].v;
    formData.pilings.measurements.left1.y = workbook.Sheets[loadedSheetName][respectiveCells[5][4]].v;
    formData.pilings.measurements.left1.w = workbook.Sheets[loadedSheetName][respectiveCells[5][5]].v;
    formData.pilings.measurements.left2.x = workbook.Sheets[loadedSheetName][respectiveCells[5][6]].v;
    formData.pilings.measurements.left2.y = workbook.Sheets[loadedSheetName][respectiveCells[5][7]].v;
    formData.pilings.measurements.left2.w = workbook.Sheets[loadedSheetName][respectiveCells[5][8]].v;
    formData.pilings.measurements.left3.x = workbook.Sheets[loadedSheetName][respectiveCells[5][9]].v;
    formData.pilings.measurements.left3.y = workbook.Sheets[loadedSheetName][respectiveCells[5][10]].v;
    formData.pilings.measurements.left3.w = workbook.Sheets[loadedSheetName][respectiveCells[5][11]].v;
    formData.pilings.measurements.left4.x = workbook.Sheets[loadedSheetName][respectiveCells[5][12]].v;
    formData.pilings.measurements.left4.y = workbook.Sheets[loadedSheetName][respectiveCells[5][13]].v;
    formData.pilings.measurements.left4.w = workbook.Sheets[loadedSheetName][respectiveCells[5][14]].v;
    formData.pilings.measurements.left5.x = workbook.Sheets[loadedSheetName][respectiveCells[5][15]].v;
    formData.pilings.measurements.left5.y = workbook.Sheets[loadedSheetName][respectiveCells[5][16]].v;
    formData.pilings.measurements.left5.w = workbook.Sheets[loadedSheetName][respectiveCells[5][17]].v;
    formData.pilings.measurements.left6.x = workbook.Sheets[loadedSheetName][respectiveCells[5][18]].v;
    formData.pilings.measurements.left6.y = workbook.Sheets[loadedSheetName][respectiveCells[5][19]].v;
    formData.pilings.measurements.left6.w = workbook.Sheets[loadedSheetName][respectiveCells[5][20]].v;
    formData.pilings.measurements.right1.x = workbook.Sheets[loadedSheetName][respectiveCells[5][21]].v;
    formData.pilings.measurements.right1.y = workbook.Sheets[loadedSheetName][respectiveCells[5][22]].v;
    formData.pilings.measurements.right1.w = workbook.Sheets[loadedSheetName][respectiveCells[5][23]].v;
    formData.pilings.measurements.right2.x = workbook.Sheets[loadedSheetName][respectiveCells[5][24]].v;
    formData.pilings.measurements.right2.y = workbook.Sheets[loadedSheetName][respectiveCells[5][25]].v;
    formData.pilings.measurements.right2.w = workbook.Sheets[loadedSheetName][respectiveCells[5][26]].v;
    formData.pilings.measurements.right3.x = workbook.Sheets[loadedSheetName][respectiveCells[5][27]].v;
    formData.pilings.measurements.right3.y = workbook.Sheets[loadedSheetName][respectiveCells[5][28]].v;
    formData.pilings.measurements.right3.w = workbook.Sheets[loadedSheetName][respectiveCells[5][29]].v;
    formData.pilings.measurements.right4.x = workbook.Sheets[loadedSheetName][respectiveCells[5][30]].v;
    formData.pilings.measurements.right4.y = workbook.Sheets[loadedSheetName][respectiveCells[5][31]].v;
    formData.pilings.measurements.right4.w = workbook.Sheets[loadedSheetName][respectiveCells[5][32]].v;
    formData.pilings.measurements.right5.x = workbook.Sheets[loadedSheetName][respectiveCells[5][33]].v;
    formData.pilings.measurements.right5.y = workbook.Sheets[loadedSheetName][respectiveCells[5][34]].v;
    formData.pilings.measurements.right5.w = workbook.Sheets[loadedSheetName][respectiveCells[5][35]].v;
    formData.pilings.measurements.right6.x = workbook.Sheets[loadedSheetName][respectiveCells[5][36]].v;
    formData.pilings.measurements.right6.y = workbook.Sheets[loadedSheetName][respectiveCells[5][37]].v;
    formData.pilings.measurements.right6.w = workbook.Sheets[loadedSheetName][respectiveCells[5][38]].v;
    formData.pilings.measurements.itoi1 = workbook.Sheets[loadedSheetName][respectiveCells[5][39]].v;
    formData.pilings.measurements.itoi2 = workbook.Sheets[loadedSheetName][respectiveCells[5][40]].v;
    formData.pilings.measurements.itoi3 = workbook.Sheets[loadedSheetName][respectiveCells[5][41]].v;
    formData.pilings.measurements.itoi4 = workbook.Sheets[loadedSheetName][respectiveCells[5][42]].v;
    formData.pilings.measurements.itoi5 = workbook.Sheets[loadedSheetName][respectiveCells[5][43]].v;
    formData.pilings.measurements.itoi6 = workbook.Sheets[loadedSheetName][respectiveCells[5][44]].v;
    formData.pilings.cover.height = workbook.Sheets[loadedSheetName][respectiveCells[5][45]].v;
    formData.pilings.cover.width = workbook.Sheets[loadedSheetName][respectiveCells[5][46]].v;
    formData.pilings.cover.quantity = workbook.Sheets[loadedSheetName][respectiveCells[5][47]].v;
    formData.pilings.cover.type = workbook.Sheets[loadedSheetName][respectiveCells[5][48]].v;
    formData.pilings.cover.color = workbook.Sheets[loadedSheetName][respectiveCells[5][49]].v;

    // Brackets
    formData.brackets.reaches[0].ty = workbook.Sheets[loadedSheetName][respectiveCells[6][0]].v;
    formData.brackets.reaches[0].values[0] = workbook.Sheets[loadedSheetName][respectiveCells[6][1]].v;
    formData.brackets.reaches[0].values[1] = workbook.Sheets[loadedSheetName][respectiveCells[6][2]].v;
    formData.brackets.reaches[0].values[2] = workbook.Sheets[loadedSheetName][respectiveCells[6][3]].v;
    formData.brackets.reaches[0].values[3] = workbook.Sheets[loadedSheetName][respectiveCells[6][4]].v;
    formData.brackets.reaches[0].values[4] = workbook.Sheets[loadedSheetName][respectiveCells[6][5]].v;
    formData.brackets.reaches[0].values[5] = workbook.Sheets[loadedSheetName][respectiveCells[6][6]].v;
    formData.brackets.reaches[1].ty = workbook.Sheets[loadedSheetName][respectiveCells[6][7]].v;
    formData.brackets.reaches[1].values[0] = workbook.Sheets[loadedSheetName][respectiveCells[6][8]].v;
    formData.brackets.reaches[1].values[1] = workbook.Sheets[loadedSheetName][respectiveCells[6][9]].v;
    formData.brackets.reaches[1].values[2] = workbook.Sheets[loadedSheetName][respectiveCells[6][10]].v;
    formData.brackets.reaches[1].values[3] = workbook.Sheets[loadedSheetName][respectiveCells[6][11]].v;
    formData.brackets.reaches[1].values[4] = workbook.Sheets[loadedSheetName][respectiveCells[6][12]].v;
    formData.brackets.reaches[1].values[5] = workbook.Sheets[loadedSheetName][respectiveCells[6][13]].v;
    formData.brackets.reaches[2].ty = workbook.Sheets[loadedSheetName][respectiveCells[6][14]].v;
    formData.brackets.reaches[2].values[0] = workbook.Sheets[loadedSheetName][respectiveCells[6][15]].v;
    formData.brackets.reaches[2].values[1] = workbook.Sheets[loadedSheetName][respectiveCells[6][16]].v;
    formData.brackets.reaches[2].values[2] = workbook.Sheets[loadedSheetName][respectiveCells[6][17]].v;
    formData.brackets.reaches[2].values[3] = workbook.Sheets[loadedSheetName][respectiveCells[6][18]].v;
    formData.brackets.reaches[2].values[4] = workbook.Sheets[loadedSheetName][respectiveCells[6][19]].v;
    formData.brackets.reaches[2].values[5] = workbook.Sheets[loadedSheetName][respectiveCells[6][20]].v;
    formData.brackets.heights[0].ty = workbook.Sheets[loadedSheetName][respectiveCells[6][21]].v;
    formData.brackets.heights[0].values[0] = workbook.Sheets[loadedSheetName][respectiveCells[6][22]].v;
    formData.brackets.heights[0].values[1] = workbook.Sheets[loadedSheetName][respectiveCells[6][23]].v;
    formData.brackets.heights[0].values[2] = workbook.Sheets[loadedSheetName][respectiveCells[6][24]].v;
    formData.brackets.heights[0].values[3] = workbook.Sheets[loadedSheetName][respectiveCells[6][25]].v;
    formData.brackets.heights[0].values[4] = workbook.Sheets[loadedSheetName][respectiveCells[6][26]].v;
    formData.brackets.heights[0].values[5] = workbook.Sheets[loadedSheetName][respectiveCells[6][27]].v;
    formData.brackets.heights[1].ty = workbook.Sheets[loadedSheetName][respectiveCells[6][28]].v;
    formData.brackets.heights[1].values[0] = workbook.Sheets[loadedSheetName][respectiveCells[6][29]].v;
    formData.brackets.heights[1].values[1] = workbook.Sheets[loadedSheetName][respectiveCells[6][30]].v;
    formData.brackets.heights[1].values[2] = workbook.Sheets[loadedSheetName][respectiveCells[6][31]].v;
    formData.brackets.heights[1].values[3] = workbook.Sheets[loadedSheetName][respectiveCells[6][32]].v;
    formData.brackets.heights[1].values[4] = workbook.Sheets[loadedSheetName][respectiveCells[6][33]].v;
    formData.brackets.heights[1].values[5] = workbook.Sheets[loadedSheetName][respectiveCells[6][34]].v;
    formData.brackets.hardware = workbook.Sheets[loadedSheetName][respectiveCells[6][35]].v;
    formData.brackets.attachment = workbook.Sheets[loadedSheetName][respectiveCells[6][36]].v;
    formData.brackets.notes = workbook.Sheets[loadedSheetName][respectiveCells[6][37]].v;
    
    // Custom
    formData.custom.boot.l = workbook.Sheets[loadedSheetName][respectiveCells[7][0]].v;
    formData.custom.boot.w = workbook.Sheets[loadedSheetName][respectiveCells[7][1]].v;
    formData.custom.boot.h = workbook.Sheets[loadedSheetName][respectiveCells[7][2]].v;
    formData.custom.boot.x = workbook.Sheets[loadedSheetName][respectiveCells[7][3]].v;
    formData.custom.boot.y = workbook.Sheets[loadedSheetName][respectiveCells[7][4]].v;
    formData.custom.boot.color = workbook.Sheets[loadedSheetName][respectiveCells[7][5]].v;
    formData.custom.nose.bend = workbook.Sheets[loadedSheetName][respectiveCells[7][6]].v;
    formData.custom.nose.width = workbook.Sheets[loadedSheetName][respectiveCells[7][7]].v;
    formData.custom.rear.bend = workbook.Sheets[loadedSheetName][respectiveCells[7][8]].v;
    formData.custom.rear.width = workbook.Sheets[loadedSheetName][respectiveCells[7][9]].v;
    formData.custom.ridge.enabled = workbook.Sheets[loadedSheetName][respectiveCells[7][10]].v;
    // formData.custom.ridge.image
    formData.custom.frame.enabled = workbook.Sheets[loadedSheetName][respectiveCells[7][11]].v;
    // formData.custom.frame.image
    formData.custom.openPassThru.l = workbook.Sheets[loadedSheetName][respectiveCells[7][12]].v;
    formData.custom.openPassThru.w = workbook.Sheets[loadedSheetName][respectiveCells[7][13]].v;
    // formData.custom.openPassThru.h = workbook.Sheets[loadedSheetName][respectiveCells[7][14]].v;
    formData.custom.openPassThru.x = workbook.Sheets[loadedSheetName][respectiveCells[7][14]].v;
    formData.custom.openPassThru.y = workbook.Sheets[loadedSheetName][respectiveCells[7][15]].v;
    formData.custom.sideRtdCable = workbook.Sheets[loadedSheetName][respectiveCells[7][16]].v;
    formData.custom.consoleWidth = workbook.Sheets[loadedSheetName][respectiveCells[7][17]].v;

    // Additional
    formData.additional.midframe = workbook.Sheets[loadedSheetName][respectiveCells[8][0]].v;
    formData.additional.remotes = workbook.Sheets[loadedSheetName][respectiveCells[8][1]].v;
    formData.additional.ridgeSupports = workbook.Sheets[loadedSheetName][respectiveCells[8][2]].v;
    formData.additional.parts[0].name = workbook.Sheets[loadedSheetName][respectiveCells[8][3]].v;
    formData.additional.parts[0].amount = workbook.Sheets[loadedSheetName][respectiveCells[8][4]].v;
    formData.additional.parts[0].cost = workbook.Sheets[loadedSheetName][respectiveCells[8][5]].v;
    formData.additional.parts[1].name = workbook.Sheets[loadedSheetName][respectiveCells[8][6]].v;
    formData.additional.parts[1].amount = workbook.Sheets[loadedSheetName][respectiveCells[8][7]].v;
    formData.additional.parts[1].cost = workbook.Sheets[loadedSheetName][respectiveCells[8][8]].v;
    formData.additional.parts[2].name = workbook.Sheets[loadedSheetName][respectiveCells[8][9]].v;
    formData.additional.parts[2].amount = workbook.Sheets[loadedSheetName][respectiveCells[8][10]].v;
    formData.additional.parts[2].cost = workbook.Sheets[loadedSheetName][respectiveCells[8][11]].v;
    formData.additional.parts[3].name = workbook.Sheets[loadedSheetName][respectiveCells[8][12]].v;
    formData.additional.parts[3].amount = workbook.Sheets[loadedSheetName][respectiveCells[8][13]].v;
    formData.additional.parts[3].cost = workbook.Sheets[loadedSheetName][respectiveCells[8][14]].v;
    formData.additional.notes = workbook.Sheets[loadedSheetName][respectiveCells[8][15]].v;

    return formData;
}