function onOpen() {
  var ui = SpreadsheetApp.getUi();           // Uiクラスを取得する
  var menu = ui.createMenu('カレンダー同期システム');  // Uiクラスからメニューを作成する
  menu.addItem('今すぐ同期', 'registerCalender');   // メニューにアイテムを追加する
  menu.addToUi();                            // メニューをUiクラスに追加する
}

function registerCalender(){
  var calender = CalendarApp.getCalendarById("c_kt8peompt8hhu32q1r8rudj1ek@group.calendar.google.com");
  var sheet = SpreadsheetApp.openById('1Yfabf1AsSvC7IPabhO_Lno9LcQfzK_L0w3SDAn2V9SU').getSheetByName('作業予定表');
  var sheetrole = SpreadsheetApp.openById('1Yfabf1AsSvC7IPabhO_Lno9LcQfzK_L0w3SDAn2V9SU').getSheetByName('作業者名簿');
  var lastRow = sheet.getRange(1, 3).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  var contents = sheet.getRange('A2:H'+lastRow).getValues();

  for(var i=0; i< contents.length; i++){
    var[role,title,name,day,at,workstatus,contens,status] = contents[i];
    if(status == true || workstatus == true){
      continue;
    }
    switch (role){
      case "実行委員":
        var send = "E";
        break;
      case "庶務":
        var send = "F";
        break;
      case "オープニング":
        var send = "G";
        break;
      case "生徒インタビュー":
       var send = "H";
        break;
      case "生徒プレゼン":
        var send = "I";
        break;
      case "積み立て自己紹介":
        var send = "J";
        break;
      case "保護者座談会":
        var send = "K";
        break;
      case "事前サポート":
        var send = "L";
        break;
      case "職員":
        var send = "M";
        break;
      case "その他":
        var send = "N";
        break;
      case "全員":
        var send = "O";
        break;
        
      default:
    }

    var roles = sheetrole.getRange(send +"2:"+ send +"20").getValues();
    var roles = Array.prototype.concat.apply([], roles);
    var roles = roles.filter(Boolean);
    console.log(roles);
    var date = new Date(day);
    console.log(contents.length);
    sheet.getRange('H' +(i+2)).setValue("true");

    let options = {
    location: "N高等学校・S高等学校 広島キャンパス",
    description: "["+role+"]"+contens+"("+ name +")",
    guests: roles.join(),
    sendInvites: false
  }
  
    calender.createAllDayEvent(title,date,options);
  }
}