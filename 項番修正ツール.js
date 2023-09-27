/*
制限事項:
・タグまたは変換後の項番を含むセル内のテキストの一部に文字色や太字などが設定されている場合、相互変換時に解除される
  -> 変換前に書式を取得するとかして頑張ればどうにかなる？
・特定のタグが出現するより前に、タグを参照できない
  -> それ用の記法を導入すればできなくはない
・自動変換された項番の前後にゼロ幅スペースがつく
  -> 手で入れた類似の番号との区別のために必要。
*/

function onOpen() {
  SpreadsheetApp.getActiveSpreadsheet()
    .addMenu("項番修正ツール", [
      { name: "タグを項番へ変換", functionName: "tagToNumber" },
      { name: "項番をタグへ戻す", functionName: "numberToTag" }
    ]);
}

function tagToNumber() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
  const values = sheet.getDataRange().getValues()

  const mappings = {}
  const nextNums = []
  values.forEach((rowValues, row) => {
    rowValues.forEach((cellValue, col) => {
      if (cellValue === '') return

      const matches = String(cellValue).matchAll(/\|(.+?)\|/g)

      let newCellValue = cellValue
      for (const [tag] of matches) {
        const level = (tag.match(/\d+\./g) || []).length + 1

        if (!mappings[tag]) {
          if (nextNums.length < level) { nextNums[level - 1] = 0 }
          nextNums[level - 1] += 1
          for (let i = level; i < nextNums.length; i++) { nextNums[i] = 0 }
          mappings[tag] = nextNums.slice(0, level).join('.')
        }

        newCellValue = newCellValue.replaceAll(tag, '\u200b' + mappings[tag] + '\u200b')
      }

      if (cellValue != newCellValue) {
        sheet.getRange(row + 1, col + 1)
          .setNote(cellValue)
          .setValue(newCellValue)
      }
    })
  })

}

function numberToTag() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
  const values = sheet.getDataRange().getValues()

  values.forEach((rowValues, row) => {
    rowValues.forEach((cellValue, col) => {
      if (cellValue === '') return

      const newCellValue = String(cellValue).replace(/\u200b((\d+\.)*\d+)\u200b/g, '|$1|')

      if (cellValue != newCellValue) {
        sheet.getRange(row + 1, col + 1)
          .setValue(newCellValue)
          .clearNote()
      }
    })
  })
}
