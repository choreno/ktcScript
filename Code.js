var g = g || {
  Month: "Jan",
  Day: "Tue",
  Year: 2019,

  Members: [
    "고용재",
    "조창범",
    "신상헌",
    "백동욱",
    "윤치홍",
    "임좌배",
    "박윤석",
    "배희성",
    "권혁제",
    "성호준",
    "안진우",
    "이준성"
  ],

  Guests: ["김기윤*", "권태경*"],

  LastMonth: "Dec",
  LastYear: 2018,
  LastDay: "Tue",
  LastMonthSheetID: "13vlZj3vLXlnG5-glygoLnrnFlXvmPvaHQd3W0ZY27Rk",

  NamedRange_Month: "MonthRange",
  NamedRange_Result: "ResultRange",
  NamedRange_LastMonth: "LastMonthRange",
  NamedRange_Update: "UpdateRange",
  NamedRange_SparkLine: "SparkLineRange"
};

function onRun() {
  MonthSheetBuilder();

  ResultSheetBuilder();

  LastMonthSheetBuilder();

  UpdateSheetBuilder();

  SeedSheetBuilder();
}
