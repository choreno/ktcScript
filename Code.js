//This is not a javascript file, but google script

var g = g || {

  Month: 'Feb',
  Day: 'Tue',
  Year: 2019,

  Members: [
      '고용재', '조창범', '신상헌', '백동욱',
      '윤치홍', '임좌배', '박윤석', '배희성',
      '권혁제', '안진우', '김효중', '강빌립',
      '임지석', '한동우'

  ],

  Guests: ['권태경*'],


  LastMonth: 'Jan',
  LastYear: 2019,
  LastDay: 'Tue',
  LastMonthSheetID: "1BDO6ZWWn6ubo4gdjAe4nn0F2oxqf7qlwwdnwppv1oyA",

  NamedRange_Month: 'MonthRange',
  NamedRange_Result: 'ResultRange',
  NamedRange_LastMonth: 'LastMonthRange',
  NamedRange_Update: 'UpdateRange',
  NamedRange_SparkLine: 'SparkLineRange',

}


function onRun() {

  MonthSheetBuilder();

  ResultSheetBuilder();

  LastMonthSheetBuilder();

  UpdateSheetBuilder();

  SeedSheetBuilder();
}





