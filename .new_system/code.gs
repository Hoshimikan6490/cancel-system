function onFormSubmit(e) {
 //　手動初期設定
   // カレンダーIDを指定
     let Calendar_ID = "○○〇@group.calendar.google.com";
   // GoogleカレンダーのURLを指定
     let Calendar_URL = "https://calendar.google.com/calendar/~~~";
   //　データベース用のスプレッドシートのURLを設定
     let Database_spr_url = "https://docs.google.com/spreadsheets/d/~~~";
   // 登録用スプレッドシートの統計データが記録されたシートの名前を設定
     let Database_sheet = "○○○○"
 // 手動初期設定完了


 // 自動初期設定
   // タイムスタンプを変数「TimeStamp」に代入
     let TimeStamp = e.values[0];
   // ４桁番号を変数「CLNO」に代入
     let CLNO = e.values[1];
   // 氏名を変数「Name」に代入
     let Name = e.values[2];
   // メールアドレスを変数「Email」に代入
     let Email = e.values[3];
   // 予約日を変数「Cancel_day」に代入
     let Cancel_day = e.values[4];
   // 予約番号を変数「Cancel_NO」に代入
     let Cancel_NO = e.values[5];

   // 登録時のデータを取得
     // 登録用スプレッドシートの「統計データ」という名前のシートを取得
       let Yoyaku_Spreadsheet = SpreadsheetApp.openByUrl(Database_spr_url)
       let Yoyaku_Sheet = Yoyaku_Spreadsheet.getSheetByName(Database_sheet);
     // 指定された予約番号に当てはまる統計データシートの予約番号を変数「database_Cancel_NO」に代入
       let database_Cancel_NO = Yoyaku_Sheet.getRange(Cancel_NO,1).getValue()
     // 指定された予約番号に当てはまる統計データシートのメールアドレスを変数「Database_cancel_Email」に代入
       let Database_cancel_Email = Yoyaku_Sheet.getRange(Cancel_NO,4).getValue()
     // 指定された予約番号に当てはまる統計データシートの予約順を変数「Database_cancel_Ranking」に代入
       let Database_cancel_Ranking = Yoyaku_Sheet.getRange(Cancel_NO,5).getValue()
     //ログを書く
       console.log("登録用Formsに入っている予約番号：" + database_Cancel_NO)
       console.log("登録用Formsに入っているメールアドレス："+Database_cancel_Email)
       console.log("キャンセルされた予約の予約順：" + Database_cancel_Ranking)

   // エラーメッセージ用の変数作成
     let Error_reason = "error_reason"

   // ログを書く
     console.log("変数設定完了");

   // カレンダーオブジェクトを取得(赤い文字がカレンダーID。これは予定を入れる先のカレンダーIDを手入力)
   //  ※「Calendar」が実行されたときに動く内容を設定
     let Calendar = CalendarApp.getCalendarById(Calendar_ID);

   // タイムゾーン設定
     Calendar.setTimeZone("Asia/Tokyo");

   // ログを書く
     console.log("初期設定完了");
 // 自動初期設定終了


 // OKパターン
   let Can_Cleaning = function(Name) {
     let Find_events = 　Calendar.getEventsForDay(new Date(Cancel_day), {search: Database_cancel_Ranking})
     // もし、キャンセル日に予約者の名前があったときに、
       if (Find_events.length) {
         for (const event of Find_events) {
           //イベントタイトルにRankingの数字が含まれていれば、削除
             const eventTitle = event.getTitle();
             if(eventTitle.indexOf(Database_cancel_Ranking) !== -1){
               event.deleteEvent();
              }//if
          }//for

         // ログを書く
           console.log(Name + "さんの" + Cancel_day + "の" + Database_cancel_Ranking + "番の予約を削除完了")

         //　自動返信メール
           // HTMLを取得
             doGet_success

           // 取得したHTMLに変数を代入
             let html = HtmlService.createHtmlOutputFromFile("success").getContent();
               // 取得したHTMLの一番最初の「Name」を変数の「Name」の中身に置き換え（以降同様）
                 let html2 = html.replace("Name", Name);
                 let html3 = html2.replace("Cancel_day", Cancel_day);
                 let html4 = html3.replace("Calendar_URL", Calendar_URL);
                 let html5 = html4.replace("CLNO", CLNO);
                 let html6 = html5.replace("Name", Name);
                 let html7 = html6.replace("Cancel_NO", Cancel_NO);
                 let html8 = html7.replace("Cancel_day", Cancel_day);
                 let html_perfect = html8.replace("TimeStamp", TimeStamp);
           // ログを書く
             console.log(html_perfect)
           // 件名を変数「Subject」に代入
             let Subject ="【" + Name + "様へ】　削除完了しました";
           // 本文を変数「Body」に代入
             let Body = "";
           // 追加設定
             let Options ={
               // 送信者を「3Dプリンター　管理者」にする
               "name": "3Dプリンター　管理者",
               // メールの本文に変数代入の終わったHTMLデータを挿入
               "htmlBody": html_perfect
              };
           // メール送信
             MailApp.sendEmail(Email,Subject,Body,Options);
           // ログを書く
             console.log("完了メール送信完了")
         } else {
           //予定がなかった時に
             let Error_reason = "指定された予約は削除済みであったため、削除できませんでした。"
             Can_Not_Cleaning(Error_reason);
         }
     // functionの終了
       return Name;
   }


 // 失敗時のメール送信
   let Can_Not_Cleaning = function(Error_reason) {
     //　自動返信メール
       // HTMLを取得
         doGet_canot();
       // 取得したHTMLに変数を代入
         let html = HtmlService.createHtmlOutputFromFile("canot").getContent();
           // 取得したHTMLの一番最初の「Name」を変数の「Name」の中身に置き換え（以降同様）
             let html2 = html.replace("Name", Name);
             let html3 = html2.replace("Cancel_day", Cancel_day);
             let html4 = html3.replace("Error_reason", Error_reason);
             let html5 = html4.replace("Calendar_URL", Calendar_URL);
             let html6 = html5.replace("CLNO", CLNO);
             let html7 = html6.replace("Name", Name);
             let html8 = html7.replace("Cancel_day", Cancel_day);
             let html_perfect = html8.replace("TimeStamp", TimeStamp);
       // ログを書く
         console.log(html_perfect)
       // 件名を変数「Subject」に代入
         let Subject ="【" + Name + "様へ】　削除失敗！！";
       // 本文を変数「Body」に代入
         let Body = "";
       // 追加設定
         let Options ={
           // 送信者を「3Dプリンター　管理者」にする
           "name": "3Dプリンター　管理者",
           // メールの本文に変数代入の終わったHTMLデータを挿入
           "htmlBody": html_perfect
          };
       // メール送信
         MailApp.sendEmail(Email,Subject,Body,Options);
       // ログを書く
         console.log("削除エラーメールを送信完了")
     // functionの終了
       return Error_reason;
    }


 //メインプロセス
   // もし、キャンセル番号と登録番号が同じなら、
     if (Number(Cancel_NO) === Number(database_Cancel_NO)) {
     // キャンセル時のメールアドレスと登録時のメールアドレスが同じなら、
       if (Email === Database_cancel_Email) {
       // OKパターン
         Can_Cleaning(Name);
         console.log("削除完了")
       } else {
         // メールアドレスが違ったパターン
         let Error_reason = "メールアドレスが予約時と異なります。"
         Can_Not_Cleaning(Error_reason);
         console.log("予約削除失敗")
       }
     } else {
       // 予約番号が違ったパターン
       let Error_reason = "あなたのメールアドレスでその予約は行われていません。"
       Can_Not_Cleaning(Error_reason);
       console.log("予約削除失敗")
     }
 // メインプロセス終了
}


function doGet_success() {
 // 「doGet_success」が実行されたときに動く内容を設定
   var t = HtmlService.createTemplateFromFile('success.html');
   return t.evaluate();
}

function doGet_canot() {
 // 「doGet_canot」が実行されたときに動く内容を設定
   var t = HtmlService.createTemplateFromFile('canot.html');
   return t.evaluate();
}
