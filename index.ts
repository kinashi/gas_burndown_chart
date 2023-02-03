// 使うデータの型だけ定義
type Result = {
  properties: {
    Assign: {
      people?: {
        person: {
          email: string
        }
      }[]
    }
    'Story Point': {
      number: number
    }
    Status: {
      select?: {
        name: 'Not Started' | 'Progress' | 'Review' | 'QA' | 'Completed'
      }
    }
  }
}

type Data = {
  results: Result[]
}

const TYPE_RANGE = 'B2'
const HOLIDAY_COLUMN = 'D'
const IDEAL_COLUMN = 'E'
const ACTUAL_COLUMN = 'F'
const START_ROW = 6

// スプレッドシートで TYPE_RANGE にシート名を出力している
const getActiveSheetName = () => SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName()

const getActiveSheetType = (sheetName: string) => {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName)
  if (!sheet) return

  const value = sheet.getRange(TYPE_RANGE).getValue()

  return value === 'スプリント' ? 'sprint' : value === 'エピック' ? 'epic' : undefined
}

const fetchData = (option?: { filter: object }) => {
  const token = PropertiesService.getScriptProperties().getProperty('TOKEN')
  const databaseId = PropertiesService.getScriptProperties().getProperty('DATABASE_ID')
  const payload = option?.filter
    ? JSON.stringify({
        filter: option.filter,
      })
    : undefined

  const options = {
    headers: {
      Authorization: `Bearer ${token}`,
      'Notion-Version': '2022-06-28',
    },
    contentType: 'application/json',
    method: 'post' as const,
    payload,
  }

  const data = JSON.parse(
    UrlFetchApp.fetch(`https://api.notion.com/v1/databases/${databaseId}/query`, options).getContentText(),
  ) as Data

  const members = PropertiesService.getScriptProperties().getProperty('MEMBERS')?.split(',') ?? []

  return data.results.filter((result) =>
    result.properties.Assign.people?.some((people) => members.includes(people.person.email)),
  )
}

const getPoint = (data: ReturnType<typeof fetchData>) =>
  data.reduce(
    (result, { properties }) => {
      const point = properties['Story Point'].number ?? 0
      const status = properties.Status.select?.name
      const completed = status === 'Completed' || status === 'QA' ? point : 0

      return {
        all: result.all + point,
        completed: result.completed + completed,
      }
    },
    {
      all: 0,
      completed: 0,
    },
  )

const genFetchOption = ({
  type,
  targetName,
}: {
  type: 'sprint' | 'epic'
  targetName: string // スプリント名 or エピック名
}) => {
  switch (type) {
    case 'sprint': {
      return {
        filter: {
          property: 'Sprint',
          select: {
            equals: targetName,
          },
        },
      }
    }
    case 'epic': {
      return {
        filter: {
          property: 'Project',
          multi_select: {
            contains: targetName,
          },
        },
      }
    }
  }
}

// 現在のスプリント、エピックの場合のみ値を返却する
const isCurrent = ({
  sheetName,
}: {
  sheetName?: string // スプリント名 or エピック名
}) => {
  if (!sheetName) return false

  const currentSprintName = PropertiesService.getScriptProperties().getProperty('CURRENT_SPRINT_NAME')
  const currentEpicNames = PropertiesService.getScriptProperties().getProperty('CURRENT_EPIC_NAMES')?.split(',') ?? []

  return [currentSprintName, ...currentEpicNames].includes(sheetName)
}

const showNoTargetError = () => {
  const ui = SpreadsheetApp.getUi()
  ui.alert('対象のスプリント、またはエピックではありません')
}

const init = () => {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()

  if (!isCurrent({ sheetName: sheet.getName() })) {
    showNoTargetError()
    return
  }

  const type = getActiveSheetType(sheet.getName())
  if (!type) {
    showNoTargetError()
    return
  }

  const option = genFetchOption({ type, targetName: sheet.getName() })
  const data = fetchData(option)
  const { all: allPoint, completed: completedPoint } = getPoint(data)
  const initPoint = allPoint - completedPoint

  // スプリント初めのSPを入力
  sheet.getRange(`${IDEAL_COLUMN}${START_ROW}:${ACTUAL_COLUMN}${START_ROW}`).setValues([[initPoint, initPoint]])

  const lastRow = sheet.getLastRow()
  const holidaysCont = sheet
    .getRange(`${HOLIDAY_COLUMN}${START_ROW}:${HOLIDAY_COLUMN}${lastRow}`)
    .getValues()
    .flat()
    .filter((value) => !!value).length
  const workDaysCount = lastRow - START_ROW - holidaysCont

  for (let row = START_ROW + 1; row <= lastRow; row++) {
    const isHoliday = sheet.getRange(`${HOLIDAY_COLUMN}${row}`).getValue() !== ''
    const dayBeforeValue = sheet.getRange(`${IDEAL_COLUMN}${row - 1}`).getValue()

    // 休日の次の日は理想線のポイントを据え置き
    const value = isHoliday ? dayBeforeValue : dayBeforeValue - initPoint / workDaysCount

    sheet.getRange(`${IDEAL_COLUMN}${row}`).setValue(value)
  }
}

const setActualValue = (sheet: ReturnType<typeof SpreadsheetApp.getActiveSheet>, isManual = false) => {
  if (!isCurrent({ sheetName: sheet.getName() })) {
    if (isManual) showNoTargetError()
    return
  }

  const type = getActiveSheetType(sheet.getName())
  if (!type) {
    if (isManual) showNoTargetError()
    return
  }

  const option = genFetchOption({ type, targetName: sheet.getName() })
  const data = fetchData(option)
  const { all, completed } = getPoint(data)

  const lastRow =
    START_ROW + sheet.getRange(`${ACTUAL_COLUMN}${START_ROW}:${ACTUAL_COLUMN}50`).getValues().filter(String).length
  sheet.getRange(`${ACTUAL_COLUMN}${lastRow}`).setValue(all - completed)
}

const setAllActualValues = (isManual = false) => {
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets()

  sheets.forEach((sheet) => {
    setActualValue(sheet)
  })
}

const showPoint = () => {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
  const type = getActiveSheetType(sheet.getName())
  if (!type) {
    showNoTargetError()
    return
  }

  const option = genFetchOption({ type, targetName: sheet.getName() })
  const data = fetchData(option)
  const point = getPoint(data)

  SpreadsheetApp.getUi().alert(JSON.stringify(point))
}

const setActualValueManually = () => {
  const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()

  setActualValue(activeSheet, true)
}

const onOpen = () => {
  SpreadsheetApp.getActiveSpreadsheet().addMenu('バーンダウン設定', [
    { name: '初期値入力', functionName: 'init' },
    { name: '現在の実績を入力', functionName: 'setActualValueManually' },
    { name: 'ポイントを確認', functionName: 'showPoint' },
  ])
}
