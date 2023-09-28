function onOpen() {
  SpreadsheetApp.getActiveSpreadsheet()
    .addMenu("項番修正ツール", [
      { name: "タグを項番へ変換 |a.b| → a.b", functionName: "tagToNumber" },
      { name: "項番をタグへ戻す a.b → |a.b|", functionName: "numberToTag" },
      { name: "バックアップ用メモの削除", functionName: "clearNotes" },
      null,
      { name: "使用方法", functionName: "showUsage" }
    ]);
}
 
function tagToNumber() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
  const values = sheet.getDataRange().getValues()
  const notes = sheet.getDataRange().getNotes()

  const mappings = {}
  const nextNums = []
  values.forEach((rowValues, row) => {
    rowValues.forEach((cellValue, col) => {
      if (cellValue === '') return

      const matches = String(cellValue).matchAll(/\|(.+?)\|/g)

      let newCellValue = cellValue
      for (const [tag] of matches) {
        const level = (tag.match(/[^.|]+\./g) || []).length + 1

        if (!mappings[tag]) {
          if (nextNums.length < level) { nextNums[level - 1] = 0 }
          nextNums[level - 1] += 1
          for (let i = level; i < nextNums.length; i++) { nextNums[i] = 0 }
          mappings[tag] = nextNums.slice(0, level).join('.')
        }

        newCellValue = newCellValue.replaceAll(tag, '\u200b' + mappings[tag] + '\u200b')
      }

      if (cellValue != newCellValue) {
        const range = sheet.getRange(row + 1, col + 1)
        const note = notes[row][col]
        if (note === '' || note.match(/^\u200b.*\|.+\|.*\u200b$/)) {
          range.setNote('\u200b' + cellValue + '\u200b')
        }
        range.setValue(newCellValue)
      }
    })
  })

}

function numberToTag() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
  const values = sheet.getDataRange().getValues()
  const notes = sheet.getDataRange().getNotes()

  values.forEach((rowValues, row) => {
    rowValues.forEach((cellValue, col) => {
      if (cellValue === '') return

      const newCellValue = String(cellValue).replace(/\u200b((\d+\.)*\d+)\u200b/g, '|$1|')

      if (cellValue != newCellValue) {
        const range = sheet.getRange(row + 1, col + 1)
        range.setValue(newCellValue)
        if (notes[row][col].match(/^\u200b.*\|.+\|.*\u200b$/)) range.clearNote()
      }
    })
  })
}

function clearNotes() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
  const values = sheet.getDataRange().getValues()
  const notes = sheet.getDataRange().getNotes()

  values.forEach((rowValues, row) => {
    rowValues.forEach((_, col) => {
      if (notes[row][col].match(/^\u200b.*\|.+\|.*\u200b$/)) {
        sheet.getRange(row + 1, col + 1).clearNote()
      }
    })
  })
}

function showUsage() {
  Browser.msgBox(
`★ 使用方法\\n
|1.2.3| のように、項番を | (パイプ記号) で囲んでおくと、番号を追加・削除した際に自動的にずらすことができます。\\n
\\n
(例)\\n
|1| aaaaa\\n
 |1.1| xxxxx\\n
 |1.2| zzzzz\\n
|2| bbbbb\\n
 |2.1| |1.2| の結果に基づいて xxxxx\\n
\\n
のような場合に、1.1 と 1.2 の間に yyyyy の処理を追加したいとします。\\n
まず、次のように既存の項番と重複しない適当な項番(ここでは 1.1a とする)をつけて処理 yyyyy を挿入します。\\n
\\n
|1| aaaaa\\n
 |1.1| xxxxx\\n
 |1.1a| yyyyy\\n
 |1.2| zzzzz\\n
|2| bbbbb\\n
 |2.1| |1.2| の結果に基づいて xxxxx\\n
\\n
次に、「タグを項番へ変換」を押すと、以下のように項番が自動的に振り直され、パイプ記号が消えます。また、振り直された項番を参照している部分も自動的に新しい項番に更新されます。\\n
間の項番を削除した場合も数字が飛ばないように振り直されます。\\n
\\n
1 aaaaa\\n
 1.1 xxxxx\\n
 1.2 yyyyy\\n
 1.3 zzzzz\\n
2 bbbbb\\n
 2.1 1.3 の結果に基づいて xxxxx\\n
\\n
また、「項番をタグへ戻す」を押すと、自動採番された項番がパイプ記号で囲まれた状態に戻ります。\\n
\\n
|1| aaaaa\\n
 |1.1| xxxxx\\n
 |1.2| yyyyy\\n
 |1.3| zzzzz\\n
|2| bbbbb\\n
 |2.1| |1.3| の結果に基づいて xxxxx\\n
\\n
上記の手順を繰り返すことで番号を何度でも振り直すことができます。\\n
基本的に、項番がパイプ記号で囲まれた状態で編集を行い、区切りがついた際に「タグを項番へ変換」を行う使い方を想定しています。\\n
\\n
★ 細かい仕様\\n
・|1.2.3| のように、| (パイプ記号) で囲まれた、. (ピリオド) 区切りの文字列を、項番のプレースホルダ (「タグ」と呼ぶ) として認識します。
階層はピリオドの数、順序はシート左上からの出現順により判断されます (自動採番では、パイプ記号で囲まれた項番の大小には無関係に出現順に採番されます)。
ピリオド区切りであれば、数字に限らず任意の文字を使用できます。\\n
・「タグを項番へ変換」を押した際に、変換対象となったセルにはセルの変換前の内容が入ったメモがつきます (ただし、既に他のメモがついている場合は付きません)
このメモは、「バックアップ用メモの削除」を押すことで削除できます。\\n
\\n
★ 制限事項\\n
・自動採番された項番の前後にゼロ幅スペースがつきます。自動採番された項番の目印として利用しているため、ゼロ幅スペースを削除すると「項番をタグへ戻す」を実行してもパイプ記号で囲まれた状態に戻らなくなります。\\n
・タグまたは変換後の項番を含むセル内のテキストの一部に文字色や太字などの書式が設定されている場合、相互変換時に解除されます (セル全体に書式が設定されている場合は問題ありません)。\\n
・特定のタグが出現するより前に、そのタグを参照することはできません。\\n
`)
}
