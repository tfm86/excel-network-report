name: Relatório Resumo
description: ''
host: EXCEL
api_set: {}
script:
  content: |
    let indice = 1;

    $("#formatar").click(() => tryCatch(runFormatar));

    async function runFormatar() {
      await Excel.run(async (context) => {
        let sheet = context.workbook.worksheets.getItem("Base");
        sheet.getUsedRange().select();

        //Load data table
        let tempValues = sheet.getUsedRange().load(["values", "text"]);
        await context.sync();

        //Pré-Format
        let rangeALL = sheet.getRange("A1:AQ" + tempValues.values.length);
        rangeALL.format.wrapText = true;
        rangeALL.format.wrapText = false;

        //Copy sheet
        let sampleSheet = context.workbook.worksheets.getActiveWorksheet();
        let copieSheet = sampleSheet.copy(Excel.WorksheetPositionType.after, sampleSheet);

        const d = new Date();

        //Rename sheet
        let sheetCopy = context.workbook.worksheets.getItem("Base (2)");
        sheetCopy.name = "Online_" + d.getHours() + "h_" + indice;
        indice++;

        sheetCopy.activate();

        //Load data table
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

          //Load new sheet
          sheetValues.load(["values", "text"]);
          await context.sync();

          //Load new value of sheet
          myvalues = sheetValues.values;

          //format coluna tecnoliga(5)
          for (let i = 1; i < myvalues.length; i++) {
            sheetValues.getCell(i, 5).values = [[myvalues[i][6] + " - " + myvalues[i][7]]];
          }

          //Rename
          sheetValues.getCell(0, 5).values = [[myvalues[0][6]]];

          //Load new sheet
          sheetValues.load(["values", "text"]);
          await context.sync();

          //Load new value of sheet
          myvalues = sheetValues.values;

          sheetCopy.getRange("G:G").delete(Excel.DeleteShiftDirection.left);
          sheetCopy.getRange("G:G").delete(Excel.DeleteShiftDirection.left);
          await context.sync();
        }

        //Format sheetCopy
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

        //Load new sheet
        sheetValues.load(["values", "text"]);
        await context.sync();

        //Load new value of sheet
        myvalues = sheetValues.values;

        //Format Downtime, insert 1
        for (let i = 0; i < myvalues.length; i++) {
          if (
            myvalues[i][1] === "00h" ||
            myvalues[i][1] === "01h" ||
            myvalues[i][1] === "02h" ||
            myvalues[i][1] === "03h" ||
            myvalues[i][1] === "04h" ||
            myvalues[i][1] === "05h" ||
            myvalues[i][1] === "06h"
          ) {
            sheetValues.getCell(i, 32).values = [["00h - 06h"]];
            sheetValues.getCell(i, 37).values = [[1]];
          }
        }

        for (let i = 0; i < myvalues.length; i++) {
          if (
            myvalues[i][1] === "07h" ||
            myvalues[i][1] === "08h" ||
            myvalues[i][1] === "09h" ||
            myvalues[i][1] === "10h" ||
            myvalues[i][1] === "11h" ||
            myvalues[i][1] === "12h"
          ) {
            sheetValues.getCell(i, 32).values = [["06h - 12h"]];
            sheetValues.getCell(i, 38).values = [[1]];
          }
        }

        for (let i = 0; i < myvalues.length; i++) {
          if (
            myvalues[i][1] === "13h" ||
            myvalues[i][1] === "14h" ||
            myvalues[i][1] === "15h" ||
            myvalues[i][1] === "16h" ||
            myvalues[i][1] === "17h" ||
            myvalues[i][1] === "18h" ||
            myvalues[i][1] === "19h" ||
            myvalues[i][1] === "20h" ||
            myvalues[i][1] === "21h" ||
            myvalues[i][1] === "22h" ||
            myvalues[i][1] === "23h"
          ) {
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

        /*

      Format with insert 1

      */

        //Format Total, insert 1
        for (let i = 1; i < myvalues.length; i++) {
          sheetValues.getCell(i, 36).values = [[1]];
          //Format name sheet
          sheetValues.getCell(i, 41).values = [[sheetCopy.name]];
        }

        //Format Em análise, insert 1
        for (let i = 0; i < myvalues.length; i++) {
          if (!myvalues[i][5].includes("2G") && myvalues[i][18] === " / " && !myvalues[i][6].includes("_")) {
            sheetValues.getCell(i, 35).values = [[1]];
          }
        }

        //Format IMP, insert 1
        for (let i = 0; i < myvalues.length; i++) {
          if (!myvalues[i][5].includes("2G") && myvalues[i][18].includes("Intervenção")) {
            sheetValues.getCell(i, 34).values = [[1]];
          }
        }

        //Pareto, insert classification
        for (let i = 0; i < myvalues.length; i++) {
          if (!myvalues[i][5].includes("2G")) {
            if (
              myvalues[i][18].includes("Tx Terceira") ||
              myvalues[i][18].includes("Tx SWAP") ||
              myvalues[i][18].includes("Ran-sharing") ||
              myvalues[i][6].includes("_")
            ) {
              sheetValues.getCell(i, 31).values = [["PROVEDOR"]];
            }
          }
        }

        //Format 2G, insert 1
        for (let i = 0; i < myvalues.length; i++) {
          if (
            myvalues[i][5].includes("2G")
            //&& !myvalues[i][18].includes(" / ")
          ) {
            sheetValues.getCell(i, 33).values = [[1]];
            // //Zero análise
            //sheetValues.getCell(i, 31).values = [[]];
          }
        }

        sheetCopy.getRange("AF1:AK1").select();
        sheetCopy.getRange("A1").format.autofitColumns();
        //sheetCopy.getRange("G1").format.autofitColumns();
        //sheetCopy.getRange("S1").format.autofitColumns();
        sheetCopy.getRange("AF1:AK1").format.autofitColumns();
        sheetCopy.getRange("A1:AK1").format.font.bold = true;
        sheetCopy.getRange("A1:AE1").format.fill.color = "#FFFF00";
        sheetCopy.getRange("AF1:AK1").format.fill.color = "#4472C4";

        sheetCopy.getRange("AL1:AP1").format.autofitColumns();
        sheetCopy.getRange("AL1:AP1").format.font.bold = true;
        sheetCopy.getRange("AL1:AP1").format.fill.color = "#8A2BE2";

        sheetCopy.freezePanes.freezeRows();

        //Format flag 0, delete
        // Loop reverso para deletar sem pular linhas
        for (let i = myvalues.length - 1; i >= 0; i--) {
          if (myvalues[i][22] === 0) {
            //console.log(myvalues[i][10])
            sheetCopy.getRange(`${i + 1}:${i + 1}`).delete(Excel.DeleteShiftDirection.up);
          }
        }
      });
    }

    /*

    Atualizar

    */

    $("#atualizar").click(() => tryCatch(runAtualizar));

    async function runAtualizar() {
      await Excel.run(async (context) => {
        // let sheets = context.workbook.worksheets;

        //Load items sheet
        // sheets.load("items,name");
        // await context.sync();

        //const totalAbas = sheets.items.length;
        // console.log(totalAbas);

        // for (let i = 0; i < sheets.items.length; i++){
        // console.log(sheets.items[i].name);
        // }

        // sheets.items.forEach(sheet => {
        //   console.log(sheet.name);
        // });

        //Load sheet Resumo
        let sheetResumo = context.workbook.worksheets.getItem("Resumo");
        let resumoValues = sheetResumo.getUsedRange().load(["values", "text"]);
        await context.sync();

        let myResumoValeus = resumoValues.values;

        //Load sheet Resumo Tel
        let sheetResumoTel = context.workbook.worksheets.getItem("Resumo_TEL");
        let resumoTelValues = sheetResumoTel.getUsedRange().load(["values", "text"]);
        await context.sync();

        let myResumoTelValeus = resumoTelValues.values;

        // console.log(myResumoValeus[1][1]); //C3
        // console.log(myResumoValeus[2][1]); //C4
        // console.log(myResumoValeus[3][1]); //C5
        // console.log(myResumoValeus[4][1]); //C6

        //Load abas
        const sheets = context.workbook.worksheets;
        sheets.load(["items", "name"]);
        await context.sync();

        //Guardar os nomes em um array
        const nomesDasAbas = sheets.items.map((sheet) => sheet.name).filter((nome) => nome.startsWith("Online"));

        //Ordenar alfabeticamente(A → Z)
        const collator = new Intl.Collator("pt-BR", { numeric: true });
        nomesDasAbas.sort(collator.compare);

        //Resumo por aba
        for (let i = 0; i < nomesDasAbas.length; i++) {
          let onlineTemp = nomesDasAbas[i];
          let onlineResumo = context.workbook.worksheets.getItem(onlineTemp);
          let onlineResumoValues = onlineResumo.getUsedRange().load(["values", "text"]);
          await context.sync();

          let myOnlineValeus = onlineResumoValues.values;

          resumoValues.getCell(i + 1, 1).values = [[myOnlineValeus[1][41]]];
          resumoValues.getCell(i + 1, 2).values = [[myOnlineValeus.length - 1]];

          let soma2G = 0;
          let somaImp = 0;
          let somaAnalise = 0;
          let soma06 = 0,
            soma612 = 0,
            soma1224 = 0,
            soma24 = 0;

          //Total
          for (let j = 0; j < myOnlineValeus.length; j++) {
            //sum 2G
            if (myOnlineValeus[j][33] === 1) {
              soma2G = soma2G + 1;
            }
            //sum Impl
            if (myOnlineValeus[j][34] === 1) {
              somaImp = somaImp + 1;
            }
            //sum Analise
            if (myOnlineValeus[j][35] === 1) {
              somaAnalise = somaAnalise + 1;
            }
            //sum 00 - 06h
            if (myOnlineValeus[j][37] === 1 && !myOnlineValeus[j][33] && !myOnlineValeus[j][34] && !myOnlineValeus[j][35]) {
              soma06 = soma06 + 1;
            }

            //sum 06 - 12h
            if (myOnlineValeus[j][38] === 1 && !myOnlineValeus[j][33] && !myOnlineValeus[j][34] && !myOnlineValeus[j][35]) {
              soma612 = soma612 + 1;
            }
            //sum 12 - 24h
            if (myOnlineValeus[j][39] === 1 && !myOnlineValeus[j][33] && !myOnlineValeus[j][34] && !myOnlineValeus[j][35]) {
              soma1224 = soma1224 + 1;
            }

            //sum 24h
            if (myOnlineValeus[j][40] === 1 && !myOnlineValeus[j][33] && !myOnlineValeus[j][34] && !myOnlineValeus[j][35]) {
              soma24 = soma24 + 1;
            }
          }

          resumoValues.getCell(i + 1, 3).values = [[soma2G]];
          resumoValues.getCell(i + 1, 4).values = [[somaImp]];
          resumoValues.getCell(i + 1, 5).values = [[somaAnalise]];
          resumoValues.getCell(i + 1, 6).values = [[myOnlineValeus.length - 1 - soma2G - somaImp - somaAnalise]];
          resumoValues.getCell(i + 1, 7).values = [[soma06]];
          resumoValues.getCell(i + 1, 8).values = [[soma612]];
          resumoValues.getCell(i + 1, 9).values = [[soma1224]];
          resumoValues.getCell(i + 1, 10).values = [[soma24]];

          //TEL
          resumoValues.getCell(i + 8, 1).values = [[myOnlineValeus[1][41]]];

          let somaTEL = 0,
            soma06Tel = 0,
            soma0612Tel = 0,
            soma1224Tel = 0,
            soma24Tel = 0;

          for (let j = 0; j < myOnlineValeus.length; j++) {
            if (myOnlineValeus[j][31] === "TEL") {
              somaTEL = somaTEL + 1;
            }

            if (myOnlineValeus[j][31] === "TEL" && myOnlineValeus[j][32] === "00h - 06h") {
              soma06Tel = soma06Tel + 1;
            }

            if (myOnlineValeus[j][31] === "TEL" && myOnlineValeus[j][32] === "06h - 12h") {
              soma0612Tel = soma0612Tel + 1;
            }

            if (myOnlineValeus[j][31] === "TEL" && myOnlineValeus[j][32] === "12h - 24h") {
              soma1224Tel = soma1224Tel + 1;
            }

            if (myOnlineValeus[j][31] === "TEL" && myOnlineValeus[j][32] === "24h") {
              soma24Tel = soma24Tel + 1;
            }
          }

          resumoValues.getCell(i + 8, 2).values = [[somaTEL]];
          resumoValues.getCell(i + 8, 3).values = [[soma06Tel]];
          resumoValues.getCell(i + 8, 4).values = [[soma0612Tel]];
          resumoValues.getCell(i + 8, 5).values = [[soma1224Tel]];
          resumoValues.getCell(i + 8, 6).values = [[soma24Tel]];

          //PROVEDOR
          resumoValues.getCell(i + 15, 1).values = [[myOnlineValeus[1][41]]]; //C17

          let somaProvedor = 0,
            soma06Provedor = 0,
            soma0612Provedor = 0,
            soma1224Provedor = 0,
            soma24Provedor = 0;

          for (let j = 0; j < myOnlineValeus.length; j++) {
            if (myOnlineValeus[j][31] === "PROVEDOR") {
              somaProvedor = somaProvedor + 1;
            }

            if (myOnlineValeus[j][31] === "PROVEDOR" && myOnlineValeus[j][32] === "00h - 06h") {
              soma06Provedor = soma06Provedor + 1;
            }

            if (myOnlineValeus[j][31] === "PROVEDOR" && myOnlineValeus[j][32] === "06h - 12h") {
              soma0612Provedor = soma0612Provedor + 1;
            }

            if (myOnlineValeus[j][31] === "PROVEDOR" && myOnlineValeus[j][32] === "12h - 24h") {
              soma1224Provedor = soma1224Provedor + 1;
            }

            if (myOnlineValeus[j][31] === "PROVEDOR" && myOnlineValeus[j][32] === "24h") {
              soma24Provedor = soma24Provedor + 1;
            }
          }

          resumoValues.getCell(i + 15, 2).values = [[somaProvedor]];
          resumoValues.getCell(i + 15, 3).values = [[soma06Provedor]];
          resumoValues.getCell(i + 15, 4).values = [[soma0612Provedor]];
          resumoValues.getCell(i + 15, 5).values = [[soma1224Provedor]];
          resumoValues.getCell(i + 15, 6).values = [[soma24Provedor]];

          //REGIONAL
          resumoValues.getCell(i + 22, 1).values = [[myOnlineValeus[1][41]]]; //C24

          let somaRegional = 0,
            soma06Regional = 0,
            soma0612Regional = 0,
            soma1224Regional = 0,
            soma24Regional = 0;

          for (let j = 0; j < myOnlineValeus.length; j++) {
            if (myOnlineValeus[j][31] === "REGIONAL") {
              somaRegional = somaRegional + 1;
            }

            if (myOnlineValeus[j][31] === "REGIONAL" && myOnlineValeus[j][32] === "00h - 06h") {
              soma06Regional = soma06Regional + 1;
            }

            if (myOnlineValeus[j][31] === "REGIONAL" && myOnlineValeus[j][32] === "06h - 12h") {
              soma0612Regional = soma0612Regional + 1;
            }

            if (myOnlineValeus[j][31] === "REGIONAL" && myOnlineValeus[j][32] === "12h - 24h") {
              soma1224Regional = soma1224Regional + 1;
            }

            if (myOnlineValeus[j][31] === "REGIONAL" && myOnlineValeus[j][32] === "24h") {
              soma24Regional = soma24Regional + 1;
            }
          }

          resumoValues.getCell(i + 22, 2).values = [[somaRegional]];
          resumoValues.getCell(i + 22, 3).values = [[soma06Regional]];
          resumoValues.getCell(i + 22, 4).values = [[soma0612Regional]];
          resumoValues.getCell(i + 22, 5).values = [[soma1224Regional]];
          resumoValues.getCell(i + 22, 6).values = [[soma24Regional]];

          await context.sync();
        }

        //ResumoTEL por aba
        for (let i = 0; i < nomesDasAbas.length; i++) {
          let onlineTemp = nomesDasAbas[i]; //Pegar o nome da aba
          let onlineResumo = context.workbook.worksheets.getItem(onlineTemp);
          let onlineResumoValues = onlineResumo.getUsedRange().load(["values", "text"]);
          await context.sync();

          let myOnlineValeus = onlineResumoValues.values;

          let somaAm06 = 0,
            somaAm0612 = 0,
            somaAm1224 = 0,
            somaAm24 = 0;
          let somaAp06 = 0,
            somaAp0612 = 0,
            somaAp1224 = 0,
            somaAp24 = 0;
          let somaMa06 = 0,
            somaMa0612 = 0,
            somaMa1224 = 0,
            somaMa24 = 0;
          let somaPa06 = 0,
            somaPa0612 = 0,
            somaPa1224 = 0,
            somaPa24 = 0;
          let somaRr06 = 0,
            somaRr0612 = 0,
            somaRr1224 = 0,
            somaRr24 = 0;

          for (let j = 0; j < myOnlineValeus.length; j++) {
            //AM
            if (
              myOnlineValeus[j][31] === "TEL" &&
              myOnlineValeus[j][32] === "00h - 06h" &&
              myOnlineValeus[j][14] === "AM"
            ) {
              somaAm06 = somaAm06 + 1;
            }
            if (
              myOnlineValeus[j][31] === "TEL" &&
              myOnlineValeus[j][32] === "06h - 12h" &&
              myOnlineValeus[j][14] === "AM"
            ) {
              somaAm0612 = somaAm0612 + 1;
            }
            if (
              myOnlineValeus[j][31] === "TEL" &&
              myOnlineValeus[j][32] === "12h - 24h" &&
              myOnlineValeus[j][14] === "AM"
            ) {
              somaAm1224 = somaAm1224 + 1;
            }

            if (myOnlineValeus[j][31] === "TEL" && myOnlineValeus[j][32] === "24h" && myOnlineValeus[j][14] === "AM") {
              somaAm24 = somaAm24 + 1;
            }

            //AP
            if (
              myOnlineValeus[j][31] === "TEL" &&
              myOnlineValeus[j][32] === "00h - 06h" &&
              myOnlineValeus[j][14] === "AP"
            ) {
              somaAp06 = somaAp06 + 1;
            }
            if (
              myOnlineValeus[j][31] === "TEL" &&
              myOnlineValeus[j][32] === "06h - 12h" &&
              myOnlineValeus[j][14] === "AP"
            ) {
              somaAp0612 = somaAp0612 + 1;
            }
            if (
              myOnlineValeus[j][31] === "TEL" &&
              myOnlineValeus[j][32] === "12h - 24h" &&
              myOnlineValeus[j][14] === "AP"
            ) {
              somaAp1224 = somaAp1224 + 1;
            }

            if (myOnlineValeus[j][31] === "TEL" && myOnlineValeus[j][32] === "24h" && myOnlineValeus[j][14] === "AP") {
              somaAp24 = somaAp24 + 1;
            }

            //MA
            if (
              myOnlineValeus[j][31] === "TEL" &&
              myOnlineValeus[j][32] === "00h - 06h" &&
              myOnlineValeus[j][14] === "MA"
            ) {
              somaMa06 = somaMa06 + 1;
            }
            if (
              myOnlineValeus[j][31] === "TEL" &&
              myOnlineValeus[j][32] === "06h - 12h" &&
              myOnlineValeus[j][14] === "MA"
            ) {
              somaMa0612 = somaMa0612 + 1;
            }
            if (
              myOnlineValeus[j][31] === "TEL" &&
              myOnlineValeus[j][32] === "12h - 24h" &&
              myOnlineValeus[j][14] === "MA"
            ) {
              somaMa1224 = somaMa1224 + 1;
            }

            if (myOnlineValeus[j][31] === "TEL" && myOnlineValeus[j][32] === "24h" && myOnlineValeus[j][14] === "MA") {
              somaMa24 = somaMa24 + 1;
            }

            //PA
            if (
              myOnlineValeus[j][31] === "TEL" &&
              myOnlineValeus[j][32] === "00h - 06h" &&
              myOnlineValeus[j][14] === "PA"
            ) {
              somaPa06 = somaPa06 + 1;
            }
            if (
              myOnlineValeus[j][31] === "TEL" &&
              myOnlineValeus[j][32] === "06h - 12h" &&
              myOnlineValeus[j][14] === "PA"
            ) {
              somaPa0612 = somaPa0612 + 1;
            }
            if (
              myOnlineValeus[j][31] === "TEL" &&
              myOnlineValeus[j][32] === "12h - 24h" &&
              myOnlineValeus[j][14] === "PA"
            ) {
              somaPa1224 = somaPa1224 + 1;
            }

            if (myOnlineValeus[j][31] === "TEL" && myOnlineValeus[j][32] === "24h" && myOnlineValeus[j][14] === "PA") {
              somaPa24 = somaPa24 + 1;
            }

            //RR
            if (
              myOnlineValeus[j][31] === "TEL" &&
              myOnlineValeus[j][32] === "00h - 06h" &&
              myOnlineValeus[j][14] === "RR"
            ) {
              somaRr06 = somaRr06 + 1;
            }
            if (
              myOnlineValeus[j][31] === "TEL" &&
              myOnlineValeus[j][32] === "06h - 12h" &&
              myOnlineValeus[j][14] === "RR"
            ) {
              somaRr0612 = somaRr0612 + 1;
            }
            if (
              myOnlineValeus[j][31] === "TEL" &&
              myOnlineValeus[j][32] === "12h - 24h" &&
              myOnlineValeus[j][14] === "RR"
            ) {
              somaRr1224 = somaRr1224 + 1;
            }

            if (myOnlineValeus[j][31] === "TEL" && myOnlineValeus[j][32] === "24h" && myOnlineValeus[j][14] === "RR") {
              somaRr24 = somaRr24 + 1;
            }
          }

          if (i === 0) {
            resumoTelValues.getCell(i + 1, 2).values = [[somaAm06]]; //D3
            resumoTelValues.getCell(i + 1, 3).values = [[somaAm0612]]; //E3
            resumoTelValues.getCell(i + 1, 4).values = [[somaAm1224]]; //F3
            resumoTelValues.getCell(i + 1, 5).values = [[somaAm24]]; //G3

            resumoTelValues.getCell(i + 2, 2).values = [[somaAp06]]; //D4
            resumoTelValues.getCell(i + 2, 3).values = [[somaAp0612]]; //E4
            resumoTelValues.getCell(i + 2, 4).values = [[somaAp1224]]; //F4
            resumoTelValues.getCell(i + 2, 5).values = [[somaAp24]]; //G4

            resumoTelValues.getCell(i + 3, 2).values = [[somaMa06]]; //D5
            resumoTelValues.getCell(i + 3, 3).values = [[somaMa0612]]; //E5
            resumoTelValues.getCell(i + 3, 4).values = [[somaMa1224]]; //F5
            resumoTelValues.getCell(i + 3, 5).values = [[somaMa24]]; //G5

            resumoTelValues.getCell(i + 4, 2).values = [[somaPa06]]; //D6
            resumoTelValues.getCell(i + 4, 3).values = [[somaPa0612]]; //E6
            resumoTelValues.getCell(i + 4, 4).values = [[somaPa1224]]; //F6
            resumoTelValues.getCell(i + 4, 5).values = [[somaPa24]]; //G6

            resumoTelValues.getCell(i + 5, 2).values = [[somaRr06]]; //D7
            resumoTelValues.getCell(i + 5, 3).values = [[somaRr0612]]; //E7
            resumoTelValues.getCell(i + 5, 4).values = [[somaRr1224]]; //F7
            resumoTelValues.getCell(i + 5, 5).values = [[somaRr24]]; //G7
          }

          if (i === 1) {
            resumoTelValues.getCell(i + 9, 2).values = [[somaAm06]]; //D12
            resumoTelValues.getCell(i + 9, 3).values = [[somaAm0612]]; //E12
            resumoTelValues.getCell(i + 9, 4).values = [[somaAm1224]]; //F12
            resumoTelValues.getCell(i + 9, 5).values = [[somaAm24]]; //G12

            resumoTelValues.getCell(i + 10, 2).values = [[somaAp06]]; //D13
            resumoTelValues.getCell(i + 10, 3).values = [[somaAp0612]]; //E13
            resumoTelValues.getCell(i + 10, 4).values = [[somaAp1224]]; //F13
            resumoTelValues.getCell(i + 10, 5).values = [[somaAp24]]; //G13

            resumoTelValues.getCell(i + 11, 2).values = [[somaMa06]]; //D14
            resumoTelValues.getCell(i + 11, 3).values = [[somaMa0612]]; //E14
            resumoTelValues.getCell(i + 11, 4).values = [[somaMa1224]]; //F14
            resumoTelValues.getCell(i + 11, 5).values = [[somaMa24]]; //G14

            resumoTelValues.getCell(i + 12, 2).values = [[somaPa06]]; //D15
            resumoTelValues.getCell(i + 12, 3).values = [[somaPa0612]]; //E15
            resumoTelValues.getCell(i + 12, 4).values = [[somaPa1224]]; //F15
            resumoTelValues.getCell(i + 12, 5).values = [[somaPa24]]; //G15

            resumoTelValues.getCell(i + 13, 2).values = [[somaRr06]]; //D16
            resumoTelValues.getCell(i + 13, 3).values = [[somaRr0612]]; //E16
            resumoTelValues.getCell(i + 13, 4).values = [[somaRr1224]]; //F16
            resumoTelValues.getCell(i + 13, 5).values = [[somaRr24]]; //G16
          }
          if (i === 2) {
            resumoTelValues.getCell(1, 12).values = [[somaAm06]]; //N3
            resumoTelValues.getCell(1, 13).values = [[somaAm0612]]; //O3
            resumoTelValues.getCell(1, 14).values = [[somaAm1224]]; //P3
            resumoTelValues.getCell(1, 15).values = [[somaAm24]]; //Q3

            resumoTelValues.getCell(2, 12).values = [[somaAp06]]; //N4
            resumoTelValues.getCell(2, 13).values = [[somaAp0612]]; //O4
            resumoTelValues.getCell(2, 14).values = [[somaAp1224]]; //P4
            resumoTelValues.getCell(2, 15).values = [[somaAp24]]; //Q4

            resumoTelValues.getCell(3, 12).values = [[somaMa06]]; //N5
            resumoTelValues.getCell(3, 13).values = [[somaMa0612]]; //O5
            resumoTelValues.getCell(3, 14).values = [[somaMa1224]]; //P5
            resumoTelValues.getCell(3, 15).values = [[somaMa24]]; //Q5

            resumoTelValues.getCell(4, 12).values = [[somaPa06]]; //N6
            resumoTelValues.getCell(4, 13).values = [[somaPa0612]]; //O6
            resumoTelValues.getCell(4, 14).values = [[somaPa1224]]; //P6
            resumoTelValues.getCell(4, 15).values = [[somaPa24]]; //Q6

            resumoTelValues.getCell(5, 12).values = [[somaRr06]]; //N7
            resumoTelValues.getCell(5, 13).values = [[somaRr0612]]; //O7
            resumoTelValues.getCell(5, 14).values = [[somaRr1224]]; //P7
            resumoTelValues.getCell(5, 15).values = [[somaRr24]]; //Q7
          }

          if (i === 3) {
            resumoTelValues.getCell(10, 12).values = [[somaAm06]]; //N12
            resumoTelValues.getCell(10, 13).values = [[somaAm0612]]; //O12
            resumoTelValues.getCell(10, 14).values = [[somaAm1224]]; //P12
            resumoTelValues.getCell(10, 15).values = [[somaAm24]]; //Q12

            resumoTelValues.getCell(11, 12).values = [[somaAp06]]; //N13
            resumoTelValues.getCell(11, 13).values = [[somaAp0612]]; //O13
            resumoTelValues.getCell(11, 14).values = [[somaAp1224]]; //P13
            resumoTelValues.getCell(11, 15).values = [[somaAp24]]; //Q13

            resumoTelValues.getCell(12, 12).values = [[somaMa06]]; //N14
            resumoTelValues.getCell(12, 13).values = [[somaMa0612]]; //O14
            resumoTelValues.getCell(12, 14).values = [[somaMa1224]]; //P14
            resumoTelValues.getCell(12, 15).values = [[somaMa24]]; //Q14

            resumoTelValues.getCell(13, 12).values = [[somaPa06]]; //N15
            resumoTelValues.getCell(13, 13).values = [[somaPa0612]]; //O15
            resumoTelValues.getCell(13, 14).values = [[somaPa1224]]; //P15
            resumoTelValues.getCell(13, 15).values = [[somaPa24]]; //Q15

            resumoTelValues.getCell(14, 12).values = [[somaRr06]]; //N16
            resumoTelValues.getCell(14, 13).values = [[somaRr0612]]; //O16
            resumoTelValues.getCell(14, 14).values = [[somaRr1224]]; //P16
            resumoTelValues.getCell(14, 15).values = [[somaRr24]]; //Q16
          }
        }

        await context.sync();
      });
    }

    // $("#visao").click(() => tryCatch(runVisao));

    // async function runVisao() {
    //   await Excel.run(async (context) => {

    //   });
    // }

    $("#visao").click(() => tryCatch(runVisao));

    async function runVisao() {
      await Excel.run(async (context) => {
        let sheet = context.workbook.worksheets.getItem("Online_22h_1");
        sheet.getUsedRange().select();

        //Load data table
        let sheetValues = sheet.getUsedRange().load(["values", "text"]);
        await context.sync();

        let myvalues = sheetValues.values;

        //Format flag 1, delete
        // Loop reverso para deletar sem pular linhas
        for (let i = myvalues.length - 1; i >= 0; i--) {
          if (myvalues[i][22] === 0) {
            //console.log(myvalues[i][10])
            sheet.getRange(`${i + 1}:${i + 1}`).delete(Excel.DeleteShiftDirection.up);
          }
        }
      });
    }

    /** Default helper for invoking an action and handling errors. */
    async function tryCatch(callback) {
      try {
        await callback();
      } catch (error) {
        // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
        console.error(error);
      }
    }
  language: typescript
template:
  content: "<button id=\"formatar\" class=\"ms-Button\">\r\n        <span class=\"ms-Button-label\">Formatar</span>\r\n    </button>\r\n\r\n<button id=\"atualizar\" class=\"ms-Button\">\r\n        <span class=\"ms-Button-label\">Atualizar</span>\r\n    </button>\r\n\r\n\r\n<!-- <button id=\"visao\" class=\"ms-Button\">\r\n        <span class=\"ms-Button-label\">Visao</span>\r\n    </button> -->"
  language: html
style:
  content: |-
    body {
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        font-size: 14px;
        line-height: 1.5;
        padding: 10px;
    }

    section {
        margin-bottom: 20px;
    }

    h3 {
        margin-top: 0;
        margin-bottom: 10px;
        font-size: 16px;
    }

    p {
        margin: 0 0 10px 0;
    }

    button {
        background-color: #f0f0f0;
        color: #333333;
        border: 1px solid #8a8a8a;
        padding: 8px 16px;
        font-size: 14px;
        cursor: pointer;
        border-radius: 2px;
        margin-left: 20px;
        margin-bottom: 5px;
        min-width: 80px;
        display: block;
    }

    button:hover {
        background-color: #e0e0e0;
    }

    button:active {
        background-color: #d0d0d0;
    }

    input {
        padding: 8px;
        margin: 5px 0;
        border: 1px solid #ccc;
        border-radius: 2px;
        font-size: 14px;
    }

    .header {
        text-align: center;
        background-color: #f3f2f1;
        padding: 10px;
    }
  language: css
libraries: |2
    https://appsforoffice.microsoft.com/lib/1/hosted/office.js
    @types/office-js

    office-ui-fabric-js@1.4.0/dist/css/fabric.min.css
    office-ui-fabric-js@1.4.0/dist/css/fabric.components.min.css

    core-js@2.4.1/client/core.min.js
    @types/core-js

    jquery@3.1.1
    @types/jquery@3.3.1
