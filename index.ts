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

const SPRINT_NAME_RANGE = 'B3'
const HOLIDAY_COLUMN = 'D'
const IDEAL_COLUMN = 'E'
const ACTUAL_COLUMN = 'F'
const START_ROW = 6

const fetchData = (sprintName: string) => {
  const token = PropertiesService.getScriptProperties().getProperty('TOKEN')
  const databaseId = PropertiesService.getScriptProperties().getProperty('DATABASE_ID')

  const payload = JSON.stringify({
    filter: {
      property: 'Sprint',
      select: {
        equals: sprintName,
      },
    },
  })

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
      const completed = properties.Status.select?.name === 'Completed' ? point : 0

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

const init = () => {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
  const sprintName = sheet.getRange(SPRINT_NAME_RANGE).getValue() as string
  const data = fetchData(sprintName)
  const { all: allPoint } = getPoint(data)

  // スプリント初めのSPを入力
  sheet.getRange(`${IDEAL_COLUMN}${START_ROW}:${ACTUAL_COLUMN}${START_ROW}`).setValues([[allPoint, allPoint]])

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
    const value = isHoliday ? dayBeforeValue : dayBeforeValue - allPoint / workDaysCount

    sheet.getRange(`${IDEAL_COLUMN}${row}`).setValue(value)
  }
}

const setActualValue = () => {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
  const sprintName = sheet.getRange(SPRINT_NAME_RANGE).getValue() as string
  const data = fetchData(sprintName)
  const { all, completed } = getPoint(data)

  const lastRow =
    START_ROW + sheet.getRange(`${ACTUAL_COLUMN}${START_ROW}:${ACTUAL_COLUMN}20`).getValues().filter(String).length
  sheet.getRange(`${ACTUAL_COLUMN}${lastRow}`).setValue(all - completed)
}

const showPoint = () => {
  const ui = SpreadsheetApp.getUi()
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
  const sprintName = sheet.getRange(SPRINT_NAME_RANGE).getValue() as string
  const data = fetchData(sprintName)
  const point = getPoint(data)

  ui.alert(JSON.stringify(point))
}

const onOpen = () => {
  SpreadsheetApp.getActiveSpreadsheet().addMenu('バーンダウン設定', [
    { name: '初期値入力', functionName: 'init' },
    // 手動でも実行できるようにしておく
    { name: '現在の実績を入力', functionName: 'setActualValue' },
    { name: 'ポイントを確認', functionName: 'showPoint' },
  ])
}
