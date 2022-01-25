/////////////////////////////////////////////////////////////////////////////////////////
// ★本番バージョン制作手順
//　　①Formsの内容をコピー
//　　②①のFormsからExcelを開く
//　　③Excelの列の順番を「タイムスタンプ、４桁番号、氏名、メールアドレス、予約日、補助の必要性、備考」の順番に変更する。
//　　　　※その他は備考欄よりも右側においてあればOK
//　　④拡張機能＞app scriptから、エディターを起動
//　　⑤「コード.gs」にプログラムをコピー
//　　　　※既存の内容を書き換える
//　　⑥Googleの共有カレンダーを作成
//　　⑦Googleカレンダーのタイムゾーン設定を「Asia/Tokyo」に設定
//　　⑧共有カレンダーのカレンダーIDを設定
//　　⑨Apps Scriptの設定から、「『appsscript.json』マニフェスト ファイルをエディタで表示する」をONにする。
//　　⑩「appsscript.json」とGoogleカレンダーの標準時刻を”Asia/Tokyo”にする。
//　　⑪以下の設定でトリガーを設定して、保存
// 　　　・実行する関数：onFormSubmit
// 　　　・実行するデプロイ：Head
// 　　　・イベントのソース：スプレッドシートから
// 　　　・イベントの種類：フォーム送信時
// 　　　・エラー通知：１週間おきに通知を受け取る
/////////////////////////////////////////////////////////////////////////////////////////

function onFormSubmit(e) {
 //　手動初期設定
   // カレンダーIDを指定
     let Calendar_ID = "~~~@group.calendar.google.com";

   // GoogleカレンダーのURLを指定
     let Calendar_URL = "https://calendar.google.com/calendar/~~~";

   //　データベース用のスプレッドシートのURLを設定
     let Database_spr_url = "https://docs.google.com/spreadsheets/~~~";
 
   // 登録用スプレッドシートの統計データが記録されたシートの名前を設定
     let Database_sheet = "統計データ"

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
     let cancel_NO = e.values[5];

   // 登録時のデータを取得
     // 登録用スプレッドシートの「統計データ」という名前のシートを取得
       let Yoyaku_Spreadsheet = SpreadsheetApp.openByUrl(Database_spr_url)
       let Yoyaku_Sheet = Yoyaku_Spreadsheet.getSheetByName(Database_sheet);
     // 指定された予約番号に当てはまる統計データシートの予約番号を変数「database_cancel_NO」に代入
       let database_cancel_NO = Yoyaku_Sheet.getRange(cancel_NO,1).getValue()
     // 指定された予約番号に当てはまる統計データシートのメールアドレスを変数「Database_cancel_Email」に代入
       let Database_cancel_Email = Yoyaku_Sheet.getRange(cancel_NO,4).getValue()
     // 指定された予約番号に当てはまる統計データシートの予約順を変数「Database_cancel_Ranking」に代入
       let Database_cancel_Ranking = Yoyaku_Sheet.getRange(cancel_NO,5).getValue()
     //ログを書く
       console.log("登録用Formsに入っている予約番号：" + database_cancel_NO)
       console.log("登録用Formsに入っているメールアドレス："+Database_cancel_Email)
       console.log("キャンセルされた予約の予約順：" + Database_cancel_Ranking)

   // エラーメッセージ用の変数作成
     let Error_reason = "error_reason"


   // ログを書く
     console.log("変数設定完了");

   // カレンダーオブジェクトを取得
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
           // 件名を変数「Subject」に代入
             let Subject ="【" + Name + "様へ】　削除完了しました";
   
           // 本文を変数「Body」に代入
             let Body = 
               Name + "様" +"\n"
               + "\n"
               + "3Dプリンター　管理者です。" + "\n"
               + "\n"
               + "ご連絡いただき、ありがとうございます。" + Cancel_day + "の貴方のご予約をすべて削除させていただきました。" + "\n"
               + "きちんと削除されているか確認をお願いします。"　+ "\n"
               + "\n"
               + "確認はこちら↓"+ "\n"
               + Calendar_URL + "\n"
               + "\n"
               + "---------------------------------------------"+ "\n"
               + "　ご回答された内容："+ "\n"
               + "\n"
               + "　４桁番号：" + "　" + CLNO + "\n"
               + "　　お名前：" + "　" + Name + "\n"
               + "　　キャンセルを希望された日程：" + "　"+ Cancel_day + "\n"
               + "\n"
               + "　　Formsが送信された日時："　+ "　" + TimeStamp + "\n"
               + "\n"
               + "---------------------------------------------" + "\n"
               + "＜注意事項＞"+ "\n"
               + " ※お心当たりがない場合は、メールの破棄をお願いいたします。" + "\n"
               + "　※何かご不明な点がございましたら、このメールアドレスまでご返信ください。" + "\n"
               + "\n"
               + "============================================" + "\n"
               + "3Dプリンター　管理者" + "\n"
               + "============================================";

           // メール送信
             MailApp.sendEmail(Email,Subject,Body);

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
     // 件名を変数「Subject」に代入
       let Subject ="【" + Name + "様へ】　削除失敗！！";
  
     // 本文を変数「Body」に代入
       let Body = 
         Name + "様" +"\n"
         + "\n"
         + "3Dプリンター　管理者です。" + "\n"
         + "\n"
         + "ご連絡いただき、ありがとうございます。しかし、ご連絡いただいた日程(" + Cancel_day + ")から予約を削除できませんでした。" + "\n"
         + "ご確認をお願いします。"　+ "\n"
         + "\n"
         + "削除に失敗した理由はこちら↓" + "\n"
         + Error_reason + "\n"
         + "現在の予約状況の確認はこちら↓"+ "\n"
         + Calendar_URL + "\n"
         + "\n"
         + "---------------------------------------------"+ "\n"
         + "　ご回答された内容："+ "\n"
         + "\n"
         + "　４桁番号：" + "　" + CLNO + "\n"
         + "　　お名前：" + "　" + Name + "\n"
         + "　　キャンセルを希望された日程：" + "　"+ Cancel_day + "\n"
         + "\n"
         + "　　Formsが送信された日時："　+ "　" + TimeStamp + "\n"
         + "\n"
         + "---------------------------------------------" + "\n"
         + "＜注意事項＞"+ "\n"
         + " ※お心当たりがない場合は、メールの破棄をお願いいたします。" +"\n"
         + "　※何かご不明な点がございましたら、このメールアドレスまでご返信ください。" + "\n"
         + "\n"
         + "============================================" + "\n"
         + "3Dプリンター　管理者" + "\n"
         + "============================================";
 
     // メール送信
       MailApp.sendEmail(Email,Subject,Body);

     // ログを書く
       console.log("削除エラーメールを送信完了")

     // functionの終了
      return Error_reason;
  }


 //メインプロセス
   // もし、キャンセル番号と登録番号が同じなら、
     if (Number(cancel_NO) === Number(database_cancel_NO)) {
     // キャンセル時のメールアドレスと登録時のメールアドレスが同じなら、
       if (Email === Database_cancel_Email) {
       // OKパターン
         Can_Cleaning(Name);
         console.log("削除完了")
       } else {
         let Error_reason = "メールアドレスが予約時と異なります。"
         Can_Not_Cleaning(Error_reason);
         console.log("予約削除失敗")
       }
     } else {
       // ダメパターン
       let Error_reason = "あなたのメールアドレスでその予約は行われていません。"
       Can_Not_Cleaning(Error_reason);
       console.log("予約削除失敗")
     }
 // メインプロセス終了
}
