let indice = 1;

$("#formatar").click(() => tryCatch(runFormatar));

async function runFormatar() {
  await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Base");
    sheet.getUsedRange().select();

    let tempValues = sheet.getUsedRange().load(["values", "text"]);
    await context.sync();

    let rangeALL = sheet.getRange("A1:AQ" + tempValues.values.length);
    rangeALL.format.wrapText = true;
    rangeALL.format.wrapText = false;

    let sampleSheet = context.workbook.worksheets.getActiveWorksheet();
    let copieSheet = sampleSheet.copy(Excel.WorksheetPositionType.after, sampleSheet);

    const d = new Date();

    let sheetCopy = context.workbook.worksheets.getItem("Base (2)");
    sheetCopy.name = "Online_" + d.getHours() + "h_" + indice;
    indice++;

    sheetCopy.activate();

    let sheetValues = sheetCopy.getUsedRange().load(["values", "text"]);
    await context.sync();

    let myvalues = sheetValues.values;

    if (myvalues[0].length > 31) {
      sheetCopy.getRange("F:F").insert(Excel.InsertShiftDirection.right);
      sheetCopy.getRange("U:U").delete(Excel.DeleteShiftDirection.left);
      sheetCopy.getRange("W:W").delete(Excel.DeleteShiftDirection.left);
      sheetCopy.getRange("AC:AC").delete(Excel.DeleteShiftDirection.left);
      sheetCopy.getRange("AE:AE").delete(Excel.DeleteShiftDirection.left);
      await context.sync();

      sheetValues.load(["values", "text"]);
      await context.sync();

      myvalues = sheetValues.values;

      for (let i = 1; i < myvalues.length; i++) {
        sheetValues.getCell(i, 5).values = [[myvalues[i][6] + " - " + myvalues[i][7]]];
      }

      sheetValues.getCell(0, 5).values = [[myvalues[0][6]]];

      sheetValues.load(["values", "text"]);
      await context.sync();

      myvalues = sheetValues.values;

      sheetCopy.getRange("G:G").delete(Excel.DeleteShiftDirection.left);
      sheetCopy.getRange("G:G").delete(Excel.DeleteShiftDirection.left);
      await context.sync();
    }

    sheetCopy.getRange("AF1").values = [["Pareto"]];
    sheetCopy.getRange("AG1").values = [["Downtime"]];
    sheetCopy.getRange("AH1").values = [["2G"]];
    sheetCopy.getRange("AI1").values = [["IMP"]];
    sheetCopy.getRange("AJ1").values = [["Em Análise"]];
    sheetCopy.getRange("AK1").values = [["Total"]];
    sheetCopy.getRange("AL1").values = [["00H - 06H"]];
    sheetCopy.getRange("AM1").values = [["06H - 12H"]];
    sheetCopy.getRange("AN1").values = [["12H - 24H"]];
    sheetCopy.getRange("AO1").values = [[">24"]];
    sheetCopy.getRange("AP1").values = [["Sheet"]];

    sheetValues.load(["values", "text"]);
    await context.sync();

    myvalues = sheetValues.values;

    for (let i = 0; i < myvalues.length; i++) {
      if (myvalues[i][1] === "00h" || myvalues[i][1] === "01h" || myvalues[i][1] === "02h" ||
          myvalues[i][1] === "03h" || myvalues[i][1] === "04h" || myvalues[i][1] === "05h" ||
          myvalues[i][1] === "06h") {
        sheetValues.getCell(i, 32).values = [["00h - 06h"]];
        sheetValues.getCell(i, 37).values = [[1]];
      }
    }

    for (let i = 0; i < myvalues.length; i++) {
      if (myvalues[i][1] === "07h" || myvalues[i][1] === "08h" || myvalues[i][1] === "09h" ||
          myvalues[i][1] === "10h" || myvalues[i][1] === "11h" || myvalues[i][1] === "12h") {
        sheetValues.getCell(i, 32).values = [["06h - 12h"]];
        sheetValues.getCell(i, 38).values = [[1]];
      }
    }

    for (let i = 0; i < myvalues.length; i++) {
      if (myvalues[i][1] === "13h" || myvalues[i][1] === "14h" || myvalues[i][1] === "15h" ||
          myvalues[i][1] === "16h" || myvalues[i][1] === "17h" || myvalues[i][1] === "18h" ||
          myvalues[i][1] === "19h" || myvalues[i][1] === "20h" || myvalues[i][1] === "21h" ||
          myvalues[i][1] === "22h" || myvalues[i][1] === "23h") {
        sheetValues.getCell(i, 32).values = [["12h - 24h"]];
        sheetValues.getCell(i, 39).values = [[1]];
      }
    }

    for (let i = 1; i < myvalues.length; i++) {
      if (myvalues[i][1].includes("d") || myvalues[i][1].includes("m") || myvalues[i][1].includes("a")) {
        sheetValues.getCell(i, 32).values = [["24h"]];
        sheetValues.getCell(i, 40).values = [[1]];
      }
    }

    for (let i = 1; i < myvalues.length; i++) {
      sheetValues.getCell(i, 36).values = [[1]];
      sheetValues.getCell(i, 41).values = [[sheetCopy.name]];
    }

    for (let i = 0; i < myvalues.length; i++) {
      if (!myvalues[i][5].includes("2G") && myvalues[i][18] === " / " && !myvalues[i][6].includes("_")) {
        sheetValues.getCell(i, 35).values = [[1]];
      }
    }

    for (let i = 0; i < myvalues.length; i++) {
      if (!myvalues[i][5].includes("2G") && myvalues[i][18].includes("Intervenção")) {
        sheetValues.getCell(i, 34).values = [[1]];
      }
    }

    for (let i = 0; i < myvalues.length; i++) {
      if (!myvalues[i][5].includes("2G")) {
        if (myvalues[i][18].includes("Tx Terceira") || myvalues[i][18].includes("Tx SWAP") ||
            myvalues[i][18].includes("Ran-sharing") || myvalues[i][6].includes("_")) {
          sheetValues.getCell(i, 31).values = [["PROVEDOR"]];
        }
      }
    }

    for (let i = 0; i < myvalues.length; i++) {
      if (myvalues[i][5].includes("2G")) {
        sheetValues.getCell(i, 33).values = [[1]];
      }
    }

    // Field classification
    for (let i = 0; i < myvalues.length; i++) {
      if (!myvalues[i][5].includes("2G") && myvalues[i][31] !== "PROVEDOR") {
        if (myvalues[i][18].includes("Field") || myvalues[i][18] === " / ") {
          sheetValues.getCell(i, 31).values = [["Field"]];
        }
      }
    }

    sheetCopy.getRange("AF1:AK1").select();
    sheetCopy.getRange("A1").format.autofitColumns();
    sheetCopy.getRange("AF1:AK1").format.autofitColumns();
    sheetCopy.getRange("A1:AK1").format.font.bold = true;
    sheetCopy.getRange("A1:AE1").format.fill.color = "#FFFF00";
    sheetCopy.getRange("AF1:AK1").format.fill.color = "#4472C4";

    sheetCopy.getRange("AL1:AP1").format.autofitColumns();
    sheetCopy.getRange("AL1:AP1").format.font.bold = true;
    sheetCopy.getRange("AL1:AP1").format.fill.color = "#8A2BE2";

    sheetCopy.freezePanes.freezeRows();

    for (let i = myvalues.length - 1; i >= 0; i--) {
      if (myvalues[i][22] === 0) {
        sheetCopy.getRange(`${i + 1}:${i + 1}`).delete(Excel.DeleteShiftDirection.up);
      }
    }
  });
}

$("#atualizar").click(() => tryCatch(runAtualizar));

async function runAtualizar() {
  await Excel.run(async (context) => {
    let sheetResumo = context.workbook.worksheets.getItem("Resumo");
    let resumoValues = sheetResumo.getUsedRange().load(["values", "text"]);
    await context.sync();

    let myResumoValeus = resumoValues.values;

    let sheetResumoField = context.workbook.worksheets.getItem("Resumo_Field");
    let resumoFieldValues = sheetResumoField.getUsedRange().load(["values", "text"]);
    await context.sync();

    let myResumoFieldValeus = resumoFieldValues.values;

    const sheets = context.workbook.worksheets;
    sheets.load(["items", "name"]);
    await context.sync();

    const nomesDasAbas = sheets.items.map((sheet) => sheet.name).filter((nome) => nome.startsWith("Online"));

    const collator = new Intl.Collator("pt-BR", { numeric: true });
    nomesDasAbas.sort(collator.compare);

    for (let i = 0; i < nomesDasAbas.length; i++) {
      let onlineTemp = nomesDasAbas[i];
      let onlineResumo = context.workbook.worksheets.getItem(onlineTemp);
      let onlineResumoValues = onlineResumo.getUsedRange().load(["values", "text"]);
      await context.sync();

      let myOnlineValeus = onlineResumoValues.values;

      resumoValues.getCell(i + 1, 1).values = [[myOnlineValeus[1][41]]];
      resumoValues.getCell(i + 1, 2).values = [[myOnlineValeus.length - 1]];

      let soma2G = 0, somaImp = 0, somaAnalise = 0;
      let soma06 = 0, soma612 = 0, soma1224 = 0, soma24 = 0;

      for (let j = 0; j < myOnlineValeus.length; j++) {
        if (myOnlineValeus[j][33] === 1) soma2G++;
        if (myOnlineValeus[j][34] === 1) somaImp++;
        if (myOnlineValeus[j][35] === 1) somaAnalise++;
        if (myOnlineValeus[j][37] === 1 && !myOnlineValeus[j][33] && !myOnlineValeus[j][34] && !myOnlineValeus[j][35]) soma06++;
        if (myOnlineValeus[j][38] === 1 && !myOnlineValeus[j][33] && !myOnlineValeus[j][34] && !myOnlineValeus[j][35]) soma612++;
        if (myOnlineValeus[j][39] === 1 && !myOnlineValeus[j][33] && !myOnlineValeus[j][34] && !myOnlineValeus[j][35]) soma1224++;
        if (myOnlineValeus[j][40] === 1 && !myOnlineValeus[j][33] && !myOnlineValeus[j][34] && !myOnlineValeus[j][35]) soma24++;
      }

      resumoValues.getCell(i + 1, 3).values = [[soma2G]];
      resumoValues.getCell(i + 1, 4).values = [[somaImp]];
      resumoValues.getCell(i + 1, 5).values = [[somaAnalise]];
      resumoValues.getCell(i + 1, 6).values = [[myOnlineValeus.length - 1 - soma2G - somaImp - somaAnalise]];
      resumoValues.getCell(i + 1, 7).values = [[soma06]];
      resumoValues.getCell(i + 1, 8).values = [[soma612]];
      resumoValues.getCell(i + 1, 9).values = [[soma1224]];
      resumoValues.getCell(i + 1, 10).values = [[soma24]];

      resumoValues.getCell(i + 8, 1).values = [[myOnlineValeus[1][41]]];

      let somaField = 0, soma06Field = 0, soma0612Field = 0, soma1224Field = 0, soma24Field = 0;

      for (let j = 0; j < myOnlineValeus.length; j++) {
        if (myOnlineValeus[j][31] === "Field") somaField++;
        if (myOnlineValeus[j][31] === "Field" && myOnlineValeus[j][32] === "00h - 06h") soma06Field++;
        if (myOnlineValeus[j][31] === "Field" && myOnlineValeus[j][32] === "06h - 12h") soma0612Field++;
        if (myOnlineValeus[j][31] === "Field" && myOnlineValeus[j][32] === "12h - 24h") soma1224Field++;
        if (myOnlineValeus[j][31] === "Field" && myOnlineValeus[j][32] === "24h") soma24Field++;
      }

      resumoValues.getCell(i + 8, 2).values = [[somaField]];
      resumoValues.getCell(i + 8, 3).values = [[soma06Field]];
      resumoValues.getCell(i + 8, 4).values = [[soma0612Field]];
      resumoValues.getCell(i + 8, 5).values = [[soma1224Field]];
      resumoValues.getCell(i + 8, 6).values = [[soma24Field]];

      // Regional and Provider summaries follow same pattern...
      resumoValues.getCell(i + 15, 1).values = [[myOnlineValeus[1][41]]];
      let somaProvedor = 0, soma06Provedor = 0, soma0612Provedor = 0, soma1224Provedor = 0, soma24Provedor = 0;
      for (let j = 0; j < myOnlineValeus.length; j++) {
        if (myOnlineValeus[j][31] === "PROVEDOR") somaProvedor++;
        if (myOnlineValeus[j][31] === "PROVEDOR" && myOnlineValeus[j][32] === "00h - 06h") soma06Provedor++;
        if (myOnlineValeus[j][31] === "PROVEDOR" && myOnlineValeus[j][32] === "06h - 12h") soma0612Provedor++;
        if (myOnlineValeus[j][31] === "PROVEDOR" && myOnlineValeus[j][32] === "12h - 24h") soma1224Provedor++;
        if (myOnlineValeus[j][31] === "PROVEDOR" && myOnlineValeus[j][32] === "24h") soma24Provedor++;
      }
      resumoValues.getCell(i + 15, 2).values = [[somaProvedor]];
      resumoValues.getCell(i + 15, 3).values = [[soma06Provedor]];
      resumoValues.getCell(i + 15, 4).values = [[soma0612Provedor]];
      resumoValues.getCell(i + 15, 5).values = [[soma1224Provedor]];
      resumoValues.getCell(i + 15, 6).values = [[soma24Provedor]];

      resumoValues.getCell(i + 22, 1).values = [[myOnlineValeus[1][41]]];
      let somaRegional = 0, soma06Regional = 0, soma0612Regional = 0, soma1224Regional = 0, soma24Regional = 0;
      for (let j = 0; j < myOnlineValeus.length; j++) {
        if (myOnlineValeus[j][31] === "REGIONAL") somaRegional++;
        if (myOnlineValeus[j][31] === "REGIONAL" && myOnlineValeus[j][32] === "00h - 06h") soma06Regional++;
        if (myOnlineValeus[j][31] === "REGIONAL" && myOnlineValeus[j][32] === "06h - 12h") soma0612Regional++;
        if (myOnlineValeus[j][31] === "REGIONAL" && myOnlineValeus[j][32] === "12h - 24h") soma1224Regional++;
        if (myOnlineValeus[j][31] === "REGIONAL" && myOnlineValeus[j][32] === "24h") soma24Regional++;
      }
      resumoValues.getCell(i + 22, 2).values = [[somaRegional]];
      resumoValues.getCell(i + 22, 3).values = [[soma06Regional]];
      resumoValues.getCell(i + 22, 4).values = [[soma0612Regional]];
      resumoValues.getCell(i + 22, 5).values = [[soma1224Regional]];
      resumoValues.getCell(i + 22, 6).values = [[soma24Regional]];

      await context.sync();
    }

    // Resumo_Field detail by region (AM, AP, MA, PA, RR)
    for (let i = 0; i < nomesDasAbas.length; i++) {
      let onlineTemp = nomesDasAbas[i];
      let onlineResumo = context.workbook.worksheets.getItem(onlineTemp);
      let onlineResumoValues = onlineResumo.getUsedRange().load(["values", "text"]);
      await context.sync();

      let myOnlineValeus = onlineResumoValues.values;

      let somaAm06 = 0, somaAm0612 = 0, somaAm1224 = 0, somaAm24 = 0;
      let somaAp06 = 0, somaAp0612 = 0, somaAp1224 = 0, somaAp24 = 0;
      let somaMa06 = 0, somaMa0612 = 0, somaMa1224 = 0, somaMa24 = 0;
      let somaPa06 = 0, somaPa0612 = 0, somaPa1224 = 0, somaPa24 = 0;
      let somaRr06 = 0, somaRr0612 = 0, somaRr1224 = 0, somaRr24 = 0;

      for (let j = 0; j < myOnlineValeus.length; j++) {
        const isField = myOnlineValeus[j][31] === "Field";
        const dt = myOnlineValeus[j][32];
        const region = myOnlineValeus[j][14];

        if (isField && region === "AM") {
          if (dt === "00h - 06h") somaAm06++;
          if (dt === "06h - 12h") somaAm0612++;
          if (dt === "12h - 24h") somaAm1224++;
          if (dt === "24h") somaAm24++;
        }
        if (isField && region === "AP") {
          if (dt === "00h - 06h") somaAp06++;
          if (dt === "06h - 12h") somaAp0612++;
          if (dt === "12h - 24h") somaAp1224++;
          if (dt === "24h") somaAp24++;
        }
        if (isField && region === "MA") {
          if (dt === "00h - 06h") somaMa06++;
          if (dt === "06h - 12h") somaMa0612++;
          if (dt === "12h - 24h") somaMa1224++;
          if (dt === "24h") somaMa24++;
        }
        if (isField && region === "PA") {
          if (dt === "00h - 06h") somaPa06++;
          if (dt === "06h - 12h") somaPa0612++;
          if (dt === "12h - 24h") somaPa1224++;
          if (dt === "24h") somaPa24++;
        }
        if (isField && region === "RR") {
          if (dt === "00h - 06h") somaRr06++;
          if (dt === "06h - 12h") somaRr0612++;
          if (dt === "12h - 24h") somaRr1224++;
          if (dt === "24h") somaRr24++;
        }
      }

      const offsets = [
        { row: 1, col: 2, val: somaAm06 }, { row: 1, col: 3, val: somaAm0612 }, { row: 1, col: 4, val: somaAm1224 }, { row: 1, col: 5, val: somaAm24 },
        { row: 2, col: 2, val: somaAp06 }, { row: 2, col: 3, val: somaAp0612 }, { row: 2, col: 4, val: somaAp1224 }, { row: 2, col: 5, val: somaAp24 },
        { row: 3, col: 2, val: somaMa06 }, { row: 3, col: 3, val: somaMa0612 }, { row: 3, col: 4, val: somaMa1224 }, { row: 3, col: 5, val: somaMa24 },
        { row: 4, col: 2, val: somaPa06 }, { row: 4, col: 3, val: somaPa0612 }, { row: 4, col: 4, val: somaPa1224 }, { row: 4, col: 5, val: somaPa24 },
        { row: 5, col: 2, val: somaRr06 }, { row: 5, col: 3, val: somaRr0612 }, { row: 5, col: 4, val: somaRr1224 }, { row: 5, col: 5, val: somaRr24 },
      ];

      const baseRow = i === 0 ? 0 : i === 1 ? 9 : i === 2 ? 0 : 9;
      const baseCol = i < 2 ? 0 : 10;

      offsets.forEach(({ row, col, val }) => {
        resumoFieldValues.getCell(baseRow + row, baseCol + col).values = [[val]];
      });
    }

    await context.sync();
  });
}

async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    console.error(error);
  }
}
