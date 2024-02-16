import { ARGB_LIGHT_GRAY, ARGB_WHITE, RESULTS_COLS, RESULTS_TEAM_ROW_START, STANDARD_CELL_OUTLINE, TEAM_ROW_HIGHLIGHT_PROPS, TITLE_ROW_PROPS, TOURNAMENT_TYPE_TEAMS, getFont, getFontExtended } from './sheetConstants.js'
// sheet 1 Results
const generateTitleRows = (worksheet) => {
// ROW 1
  worksheet.addRow({ A: 'BUHURT INTERNATIONAL' })

  worksheet.getCell('A1').fill = TITLE_ROW_PROPS.fill
  worksheet.getCell('A1').alignment = TITLE_ROW_PROPS.alignment
  worksheet.getCell('A1').font = getFont(16, true)
  // ROW 2
  worksheet.addRow({ A: 'Event Scoring Sheet' })

  worksheet.getCell('A2').fill = TITLE_ROW_PROPS.fill
  worksheet.getCell('A2').alignment = TITLE_ROW_PROPS.alignment
  worksheet.getCell('A2').font = getFont(12, true)

  worksheet.mergeCells('A1:AA1')
  worksheet.mergeCells('A2:AA2')
}

const generateHeaderRows = (worksheet, eventName, date, location) => {
  const row3 = worksheet.addRow({
    A: 'Name:',
    B: eventName,
    T: 'Event Date:',
    W: date,
    Y: 'Event Tier:',
    Z: 'Classic',
    AA: ':Event tier',
  })

  row3.getCell('T').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }
  row3.getCell('W').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }
  row3.getCell('Y').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }
  row3.getCell('Z').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }
  row3.getCell('AA').alignment = { vertical: 'middle', horizontal: 'left', wrapText: true }

  row3.font = getFont(11, true)
  worksheet.getCell('Z3').dataValidation = {
    type: 'list',
    allowBlank: true,
    formulae: ['"Classic,Regional,Conference"']
  }
  worksheet.mergeCells('B3:S3')
  worksheet.mergeCells('T3:V3')
  worksheet.mergeCells('W3:X3')

  const row4 = worksheet.addRow({
    A: 'Loc:',
    B: location,
    T: '',
    W: '',
    Y: 'Tier Mult.',
    Z: { formula: 'IF(Z3="Regional",1.5,IF(Z3="Conference",2,1))' },
    AA: 'Tier Multiplier',
  })

  row4.getCell('T').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }
  row4.getCell('W').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }
  row4.getCell('Y').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }
  row4.getCell('Z').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }
  row4.getCell('AA').alignment = { vertical: 'middle', horizontal: 'left', wrapText: true }

  row4.font = getFont(11, true);
  worksheet.getCell('W4').dataValidation = null;
  worksheet.getCell('W4').value = 'Round Robin';
  
  worksheet.mergeCells('B4:S4')
  worksheet.mergeCells('T4:V4')
  worksheet.mergeCells('W4:X4')
}

const generateTeamHeaders = (worksheet, tournamentType) => {
  const headerFont = getFont(11, true)
  const subHeaderFont = getFont(10, true)
  const wrapAlignment = { vertical: 'middle', horizontal: 'center', wrapText: true }

  // Add headers for row 5 with common font and alignment
  const row5 = worksheet.addRow({
    A: 'ID#',
    B: 'Team',
    C: 'T',
    D: 'Fights',
    I: 'Rounds',
    N: 'Active/Grounded',
    T: 'Cards',
    W: 'Points',
    X: 'Placement',
    Y: 'Rank adj. Points',
    Z: 'Final Awarded Points'
  })

  row5.font = headerFont
  row5.alignment = wrapAlignment

  // Add headers for row 6 and styles
  const row6 = worksheet.addRow({
    // Fights
    D: 'Won',
    E: 'Lost',
    F: 'Fought',
    G: 'Ratio',
    H: 'F/T',
    // Rounds
    I: 'Won',
    J: 'Draw',
    K: 'Lost',
    L: 'Fought',
    M: 'Ratio',
    // Active/Grounded
    N: 'A',
    O: 'A/R',
    P: 'G',
    Q: 'G/R',
    R: 'A/G D', // active versus ground difference
    S: 'A/G R', // active ground ratio read "kill/death Ratio"
    // penaties
    T: 'Yellow',
    U: 'Red',
    V: 'Total'
  })

  row6.font = subHeaderFont
  row6.alignment = wrapAlignment

  worksheet.mergeCells('A5:A6')
  worksheet.mergeCells('B5:B6')
  worksheet.mergeCells('C5:C6')
  worksheet.mergeCells('D5:H5')
  worksheet.mergeCells('I5:M5')
  worksheet.mergeCells('N5:S5')
  worksheet.mergeCells('W5:W6')
  worksheet.mergeCells('X5:X6')
  worksheet.mergeCells('Y5:Y6')
  worksheet.mergeCells('Z5:Z6')
  worksheet.mergeCells('T5:V5')

  worksheet.getColumn('H').hidden = true
  worksheet.getColumn('C').hidden = true
  worksheet.getColumn('Y').hidden = true
}

// generates the team names + id
const generateTeamDataRows = (worksheet, teams) => {
  teams.forEach((team, index) => {
    const rowIndex = index + RESULTS_TEAM_ROW_START;
    const isEvenRow = (index + 1) % 2 === 0;

    const addedRow = worksheet.addRow({
      A: index,
      B: team.name,
      C: 1,
      D: { formula: `SUMIF(pools!$B:$B,$A${rowIndex},pools!$I:$I)+SUMIF(brackets!$B:$B,$A${rowIndex},brackets!$I:$I)` },
      E: { formula: `SUMIF(pools!$B:$B,$A${rowIndex},pools!$J:$J)+SUMIF(brackets!$B:$B,$A${rowIndex},brackets!$J:$J)` },
      F: { formula: `SUM(D${rowIndex}:E${rowIndex})` },
      G: { formula: `IFERROR(D${rowIndex}/F${rowIndex}, 0)` },
      H: { formula: `F${rowIndex}/C${rowIndex}` },
      I: { formula: `SUMIF(pools!$B:$B,$A${rowIndex},pools!$K:$K)+SUMIF(brackets!$B:$B,$A${rowIndex},brackets!$K:$K)` },
      J: { formula: `SUMIF(pools!$B:$B,$A${rowIndex},pools!$L:$L)+SUMIF(brackets!$B:$B,$A${rowIndex},brackets!$L:$L)` },
      K: { formula: `SUMIF(pools!$B:$B,$A${rowIndex},pools!$M:$M)+SUMIF(brackets!$B:$B,$A${rowIndex},brackets!$M:$M)` },
      L: { formula: `SUM(I${rowIndex}:K${rowIndex})` },
      M: { formula: `IFERROR(I${rowIndex}/L${rowIndex}, 0)` },
      N: { formula: `SUMIF(pools!$B:$B,$A${rowIndex},pools!$O:$O)+SUMIF(brackets!$B:$B,$A${rowIndex},brackets!$O:$O)` },
      O: { formula: `IFERROR(N${rowIndex}/L${rowIndex}, 0)` },
      P: { formula: `SUMIF(pools!$B:$B,$A${rowIndex},pools!$P:$P)+SUMIF(brackets!$B:$B,$A${rowIndex},brackets!$P:$P)` },
      Q: { formula: `IFERROR(P${rowIndex}/L${rowIndex}, 0)` },
      R: { formula: `N${rowIndex}-P${rowIndex}` },
      S: { formula: `IFERROR(N${rowIndex}/(L${rowIndex}*5), 0)` },
      T: { formula: `SUMIF(pools!$B:$B,$A${rowIndex},pools!$R:$R)+SUMIF(brackets!$B:$B,$A${rowIndex},brackets!$R:$R)` },
      U: { formula: `SUMIF(pools!$B:$B,$A${rowIndex},pools!$S:$S)+SUMIF(brackets!$B:$B,$A${rowIndex},brackets!$S:$S)` },
      V: { formula: `T${rowIndex}+(2*U${rowIndex})` },
      W: { formula: `SUMIF(pools!$B:$B, $A${rowIndex}, pools!$I:$I) + (SUMIF(brackets!$B:$B, $A${rowIndex}, brackets!$I:$I) * 2)` },
      X: { formula: `SUMIF(pools!$B:$B, $A${rowIndex}, pools!$I:$I) + (SUMIF(brackets!$B:$B, $A${rowIndex}, brackets!$I:$I) * 2)` },
      // Continue up to column AA
      Y: { formula: `IFERROR(COUNTIFS($D$7:$D$100,">"&D${rowIndex})+COUNTIFS($D$7:$D$100,D${rowIndex},$M$7:$M$100,">"&M${rowIndex})+COUNTIFS($D$7:$D$100,D${rowIndex},$M$7:$M$100,M${rowIndex},$S$7:$S$100,">"&S${rowIndex})+COUNTIFS($D$7:$D$100,D${rowIndex},$M$7:$M$100,M${rowIndex},$S$7:$S$100,S${rowIndex},$V$7:$V$100,"<"&V${rowIndex}),IFERROR(MAX(1,SUMPRODUCT(--($A$7:$A$100=A${rowIndex}),--($B$7:$B$100=B${rowIndex}),--($C$7:$C$100=C${rowIndex}),--($D$7:$D$100=D${rowIndex}),--($M$7:$M$100>M${rowIndex}),ROW($M$7:$M$100)-MIN(ROW($M$7:$M$100))+1)),"NA"))+1` },
      Z: { formula: `IF(X${rowIndex}=3,W${rowIndex}+2,IF(X${rowIndex}=2,W${rowIndex}+4,IF(X${rowIndex}=1,W${rowIndex}+6,W${rowIndex})))` },
      // AA: { formula: `Y${rowIndex}*$Z$4` }
    });
    
    // Set alternating row colors
    const fillColor = isEvenRow ? ARGB_LIGHT_GRAY : ARGB_WHITE;
    for (let colIndex = 1; colIndex <= 26; colIndex++) {
      const cell = addedRow.getCell(colIndex);
      const dataFont = getFontExtended(11, false);
      const dataAlignment = { vertical: 'middle', horizontal: 'center', wrapText: false };
      cell.font = dataFont;
      cell.alignment = dataAlignment;
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fillColor } };
      cell.border = {
        top: { style: 'thin', color: { argb: STANDARD_CELL_OUTLINE } },
        left: { style: 'thin', color: { argb: STANDARD_CELL_OUTLINE } },
        bottom: { style: 'thin', color: { argb: STANDARD_CELL_OUTLINE } },
        right: { style: 'thin', color: { argb: STANDARD_CELL_OUTLINE } },
      };
    }
    worksheet.getCell(`G${rowIndex}`).numFmt = '0.00%';
    worksheet.getCell(`M${rowIndex}`).numFmt = '0.00%';
    worksheet.getCell(`S${rowIndex}`).numFmt = '0.00%';
  });
}

const generateResultsSheet = (sheet, { eventName, date, location, teams, tournamentType }) => {
  sheet.columns = RESULTS_COLS
  generateTitleRows(sheet)
  generateHeaderRows(sheet, eventName, date, location, teams)
  generateTeamHeaders(sheet, tournamentType)
  generateTeamDataRows(sheet, teams)
}

export default generateResultsSheet
