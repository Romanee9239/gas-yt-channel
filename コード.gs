function dairyCrowl(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var list_sht = sheet.getSheetByName("チャンネル一覧");
  var last_row = list_sht.getLastRow() - 1;
  
  /*URL一覧を取得
   * data[0]...url
   * data[1]...総再生数
   */
  var data = list_sht.getRange(2, 3, last_row, 2).getDisplayValues();
  
  /* 1日～30日前までのデータ取得
   *
   */
  var dairy = list_sht.getRange(2, 6, last_row, 29).getDisplayValues();
  
  var url;
  var response;
  var html;
  var view_STag = '> &bull; <b>';
  var view_ETag = '</b> views</span>';
  var null_STag = '視聴回数 "},{"text":"';
  var null_ETag = '","bold":true}';
  
  var totalView = [];
  var onedayView = [];
  var yesterView = 0;
  var newData = [];
  
  var k = 0;
  //スクレイピング
  for(var i = 0; i < data.length; i++){
    
    url = data[i][0] + '/about';
    
    yesterView = data[i][1].split(',').join('');
    
    //総再生数取得
    totalView[i] = null;
    
    k=0;
    while(totalView[i] == null){
      response = UrlFetchApp.fetch(url);
      html = response.getContentText('UTF-8');
      //Logger.log(i + ':' + url);
      totalView[i] = bitweenHtml(html,view_STag,view_ETag);
      if (totalView[i] == null){
        totalView[i] = bitweenHtml(html,null_STag,null_ETag);
        k++;
      }
      if (k > 500 && totalView[i] == null){
        for (var n=0; n < 10; n++){
          var document = DocumentApp.create("nullHTML" + n);
          k = html.length/10;
          document.getBody().setText(html.slice(k*n,k*(n+1)));
          document.saveAndClose();
        }
        totalView[i] = yesterView;
      }
    }
    if (k > 0){
      //Logger.log(totalView[i]);
    }
    if (totalView[i] != null){
      totalView[i] = totalView[i].split(',').join('');
    }else{
      Logger.log("55:nullやで");
    }
    
    Logger.log('38:' + data[i][1]);
    //昨日１日の再生数取得
    
    if(isNaN(yesterView)){
      yesterView = 0;
    }
    onedayView[i] = totalView[i] - yesterView;
    
    newData[i] = [totalView[i],
                  '=average(F' + (i + 2) + ':AI' + (i + 2) + ')',
                  onedayView[i]
                  ];
                  
    for (var n = 0; n < dairy[i].length; n++){
      newData[i][n + 3] = dairy[i][n].split(',').join('');
    }
    
  }
  Browser.msgBox(newData[0].length);
  list_sht.getRange(2, 4, last_row, newData[0].length).setValues(newData);
}

/*****************************
 * スクレイピング用メソッド
 * @param1 : html
 * @param2 : searchTag
 * @param3 : endTag
 * @return : スクレイピングしたやつ
 *****************************/
function bitweenHtml(html, searchTag, endTag){
  var res_val = null;
  var index = html.indexOf(searchTag);
  if (index !== -1) { 
    var html_sun = html.substring(index + searchTag.length,html.length);
    index = html_sun.indexOf(endTag);
    if (index !== -1) {
      res_val = html_sun.substring(0, index);
    }
  }
  
  return res_val;
}