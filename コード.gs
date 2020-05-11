function onOpen() {
  var book = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [
    {
      name : "チャンネルの動画データを取得",
      functionName : "getMovieData"
    }
  ];
  sheet.addMenu("スクリプト実行", entries);

};

/**********************************************
 *検索結果から動画データを取得するマン
 *
 **********************************************/
function getSearchData(){
  
  var book = SpreadsheetApp.getActiveSpreadsheet();
  var sheetData = book.getSheetByName("入力");
  
  const APIURL_GET_CHANNEL  = 'https://www.googleapis.com/youtube/v3/channels?';
  const APIURL_GET_PLAYLIST = 'https://www.googleapis.com/youtube/v3/playlistItems?';
  const APIURL_GET_VIDEODATA = 'https://www.googleapis.com/youtube/v3/videos?';
  const APIURL_GET_SEARCHRES = 'https://www.googleapis.com/youtube/v3/search??type=video';
  
  const CHANNEL_URL_HEAD_DF = 'https://www.youtube.com/channel/';
  const CHANNEL_URL_HEAD_US = 'https://www.youtube.com/user/';


  var inputCelCol = 3;
  var inputCelRow = 1;
  
  var colURL = 6;
  var colContributeCount = 12;
  var colIcon = 4;

  var rowStartData = 6
  var rowEndData = sheetData.getDataRange().getLastRow()
  

  //設定値取得
  /*
  var c = 5;
  var totalViewBuzzCond = sheetData.getRange(c++, 2).getValue();  //バズったと思う視聴回数
  var resentPeriodCond   = sheetData.getRange(c++, 2).getValue();  //初期バズ観察日数
  var periodBuzzCond    = sheetData.getRange(c++, 2).getValue();  //初期バズ判定視聴数
  */
  
  //apiキー取得
  var key = {pri : ''};
  var key_arr = [];
  
  var key_col = 6;
  var lastRow = sheetData.getRange(sheetData.getMaxRows(), key_col).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  var counta = 0;
  for(var i = 10; i <= lastRow; i++){
    key_arr[counta] = sheetData.getRange(i, key_col).getValue();
    if(key_arr[counta] != ''){ 
      //Browser.msgBox(i + ':' + key_arr[counta]);
      counta++; 
    }
  }
  key.pri = key_arr[0];
  
  //検索ワードを取得
  var word_empty = false;
  var word = '@not';
  var row = inputCelRow;
  var keyword = '';
  word = sheetData.getRange(row++,inputCelCol).getValue();
  while(word != ''){
    keyword += word;
    word = sheetData.getRange(row++,inputCelCol).getValue();
    if (word != ''){
      keyword += '+';
    }
  }
  //取得件数
  var maxResult = 20;
  
  
  var keys = "&key=" + key.pri;
  var prm = "&part=snippet&maxResults=" + maxResult;
  var order = '&order=viewCount';
  var search_word = '&q=' + keyword;
  
  var targetChannelUrl = [];
  var targetChannelName = [];
  
  var targetVideoId = [];
  var targetVideoUrl = [];
  var targetVideoTitle = [];
  var targetVideoDate = [];
  var targetOneWeek = [];
  var targetVideoTime = [];
  var targetVideoView = [];
  var targetVideoDuration = [];
  var targetLikeCount = [];
  var targetDislikeCount = [];
  var targetComCount = [];
  var targetFileName = [];
  var targetVideoTags = [];
  var ctitle = "";
  
  var q_suc = false;
  
  
  var dataURL = APIURL_GET_SEARCHRES + prm + search_word + order + keys;
  
  
  var resp;
  resp = requestApi(dataURL, key, key_arr);
  if (resp == false){ return; }
  /*
  var c_json  = JSON.parse(resp.getContentText());
  var c_items = c_json.items[0];
  var c_snippet = c_items.snippet;
  var c_statistics = c_items.statistics;
  var c_contentDetails = c_items.contentDetails;
  */
  
  var list = JSON.parse(resp.getContentText()).items;
  
  var vjson;
  var vprm;
  var vurl;
  var data;
  
end_label1:
  for(var k = 0; k< maxResult; k+=1){
    
    if(list[k] == null){
      k--;
      break end_label1;
    }
    
    //ID
    targetVideoId[k] = list[k].id.videoId;
    //url
    targetVideoUrl[k] = "https://www.youtube.com/watch?v=" + targetVideoId[k];
    //タイトル
    targetVideoTitle[k] = list[k].snippet.title;
    //日時
    var publishedTime = list[k].snippet.publishedAt;
    targetVideoDate[k] = Utilities.formatDate(new Date(publishedTime), "JST", "yyyy/MM/dd");
    targetVideoTime[k] = Utilities.formatDate(new Date(publishedTime), "JST", "HH:mm:ss");
    
    //チャンネル名・URL
    targetChannelUrl[k] = 'https://www.youtube.com/channel/' + list[k].snippet.channelId;
    targetChannelName[k] = list[k].snippet.channelTitle;
    
    vprm = "&part=snippet,contentDetails,statistics&id=" + targetVideoId[k];
    
    vurl = APIURL_GET_VIDEODATA + vprm + keys;
    
    //Browser.msgBox(vurl);
    
    vjson = requestApi(vurl, key, key_arr);
    if (vjson == false){ return; }
    data = JSON.parse(vjson.getContentText()).items[0];
    
    try{
      
      //再生数
      targetVideoView[k] = data.statistics.viewCount;
      //動画時間
      var videoDuration = data.contentDetails.duration.replace('PT','').replace('S','').split('M');
      if (videoDuration[1] == 0) {
        videoDuration[1] = "00";
      }
      else if (videoDuration[1] < 10) {
        videoDuration[1] = "0" + videoDuration[1];
      }
      targetVideoDuration[k] = videoDuration[0] + ':' + videoDuration[1];
      //高評価
      targetLikeCount[k] = data.statistics.likeCount;
      //低評価
      targetDislikeCount[k] = data.statistics.dislikeCount;
      //コメント数
      targetComCount[k] = data.statistics.commentCount;
      //タグ取得
      targetVideoTags[k] = data.snippet.tags;
      
    }catch(e){
      //再生数
      targetVideoView[k] = 0;
      //動画時間
      targetVideoDuration[k] = '0:00';
      //高評価
      targetLikeCount[k] = 0;
      //低評価
      targetDislikeCount[k] = 0;
      //コメント数
      targetComCount[k] = 0;
      //タグ取得
      targetVideoTags[k] = '-';
    }
  }
  
  order = '&order=date';
  dataURL = APIURL_GET_SEARCHRES + prm + search_word + order + key;
  resp = UrlFetchApp.fetch(dataURL);
  
  list = JSON.parse(resp.getContentText()).items;
  
  var count = k;
  
end_label2:
  for(var x = 0; x < maxResult; x+=1){
    
    if(list[x] == null){
        k--;
        break end_label2;
      }
    
    Logger.log(k + '件目：' + list[x]);
    
    if(list[x].id.videoId == undefined){
        k--;
        continue;
    }
    
    //ID
    targetVideoId[k] = list[x].id.videoId;    
    //url
    targetVideoUrl[k] = "https://www.youtube.com/watch?v=" + targetVideoId[k];
    //タイトル
    targetVideoTitle[k] = list[x].snippet.title;
    //日時
    publishedTime = list[x].snippet.publishedAt;
    targetVideoDate[k] = Utilities.formatDate(new Date(publishedTime), "JST", "yyyy/MM/dd");
    targetVideoTime[k] = Utilities.formatDate(new Date(publishedTime), "JST", "HH:mm:ss");
    
    //チャンネル名・URL
    targetChannelUrl[k] = 'https://www.youtube.com/channel/' + list[x].snippet.channelId;
    targetChannelName[k] = list[x].snippet.channelTitle;
    
    vprm = "&part=snippet,contentDetails,statistics&id=" + targetVideoId[k];
    
    vurl = APIURL_GET_VIDEODATA + vprm + keys;
    
    vjson = requestApi(APIURL_GET_VIDEODATA + vprm, key, key_arr);
    if (vjson == false){ return; }
    data = JSON.parse(vjson.getContentText()).items[0];
    
    try{
      //再生数
      targetVideoView[k] = data.statistics.viewCount;
      Logger.log(targetVideoView[k]);
      
      //動画時間
      var videoDuration = data.contentDetails.duration.replace('PT','').replace('S','').split('M');
      if (videoDuration[1] == 0) {
        videoDuration[1] = "00";
      }
      else if (videoDuration[1] < 10) {
        videoDuration[1] = "0" + videoDuration[1];
      }
      targetVideoDuration[k] = videoDuration[0] + ':' + videoDuration[1];
      //高評価
      targetLikeCount[k] = data.statistics.likeCount;
      //低評価
      targetDislikeCount[k] = data.statistics.dislikeCount;
      //コメント数
      targetComCount[k] = data.statistics.commentCount;
      //タグ取得
      targetVideoTags[k] = data.snippet.tags;
    
    }catch(e){
      //再生数
      targetVideoView[k] = 0;
      //動画時間
      targetVideoDuration[k] = '0:00';
      //高評価
      targetLikeCount[k] = 0;
      //低評価
      targetDislikeCount[k] = 0;
      //コメント数
      targetComCount[k] = 0;
      //タグ取得
      targetVideoTags[k] = '-';
    }
    
    k++;
  }
  
  //チャンネルのシート取得
  var resSheet = "結果";
  var newSheet = book.getSheetByName(resSheet);
  if(newSheet != null){
    book.deleteSheet(newSheet);
  }
  book.insertSheet(resSheet);
  newSheet = book.getSheetByName(resSheet);
  
  var defRow = 1;
  
  var m = 1
  newSheet.getRange(defRow, m++).setValue('チャンネル名');
  newSheet.getRange(defRow, m++).setValue('タイトル');
  var date = m;
  newSheet.getRange(defRow, m++).setValue('投稿日');
  newSheet.getRange(defRow, m++).setValue('曜日');
  newSheet.getRange(defRow, m++).setValue('投稿時間');
  var view = m;  
  newSheet.getRange(defRow, m++).setValue('再生数');
  newSheet.getRange(defRow, m++).setValue('日次再生');
  var h = m;
  newSheet.getRange(defRow, m++).setValue('動画時間');
  newSheet.getRange(defRow, m++).setValue('高評価数');
  newSheet.getRange(defRow, m++).setValue('低評価数');
  newSheet.getRange(defRow, m++).setValue('コメント');
  var tags = m;
  newSheet.getRange(defRow, m++).setValue('タグ');
  var tmp = m;
  
  var tempCel = newSheet.getRange(1, 3);
  tempCel.setValue(todate);
  
  //二次元配列の生成
  var targetValue = [];
  var l = 0;
  Logger.log(count);
  for(var y = 0; y < k + 1; y+=1){
    
    if (targetVideoTags[l] == null){
      //none
    }else{
      var tag = targetVideoTags[l][0];
      for(var t=1; t<targetVideoTags[l].length; t+=1){
        tag = tag + ',' + targetVideoTags[l][t];
      }
    }
    
    if(y == count){
      targetValue[y] = ['日付順','','','','','','','','','','','',''];
    }else{
      Logger.log(l + '件目：' + targetVideoTitle[l]);
      targetValue[y] = [
        '=HYPERLINK("' + targetChannelUrl[l] + '","' + targetChannelName[l] + '")',
        '=HYPERLINK("' + targetVideoUrl[l] + '","' + targetVideoTitle[l].split('"').join('""') + '")',
        targetVideoDate[l],
        '=text(C' + (l+3) + ',"ddd")',
        targetVideoTime[l],
        targetVideoView[l],
        '=F' + (l+3) + '/(today() + 1 - C' + (l+3) + ')',
        targetVideoDuration[l],
        targetLikeCount[l],
        targetDislikeCount[l],
        targetComCount[l],
        tag,
        '=C' + (l+3) + '-C1'
      ];
      l++;
      Logger.log(l + '件目');
    }
  }
  
  newSheet.getRange(defRow + 1, 1, targetValue.length,targetValue[0].length).setValues(targetValue);
  
  /*
  ハイライト
  var nn = 0;
  var videoView;
  
  for(var n = 0; n<targetValue.length; n+=1){  
  
    nn = n + o + 1;
    
    videoView = newSheet.getRange(nn, view);
    
    //特定回数以上再生された動画は再生回数をハイライトする
    if (videoView.getValue() > totalViewBuzzCond){
      videoView.setBackground('#ffff00');
      videoView.setFontStyle('Bold');
    }
    
  }  
  
  newSheet.getRange(4, tmp, nn, 1).clear();
  
  newSheet.setFrozenRows(o);
  */
  
  for(var c=1; c<= 10; c+=1){
    newSheet.autoResizeColumn(c);
  }
  
}

function getDataByCode(text,startTag,endTag){
  var index = text.indexOf(startTag);
  if (index !== -1) {
    var html_sun = text.substring(index + startTag.length,text.length);
    var index = html_sun.indexOf('endTag');
    if (index !== -1) {
      return html_sun.substring(index + endTag.length, html_sun.length);
    }
  }
  return "";
}

/**********************************************
 *APIなどkeyを取得
 *@param    keyCol...目的のKeyが書いてある列
 *@return   key配列
 **********************************************/
function getNextKey(key_arr, key){
  var res_val = 0;
  var num = key_arr.indexOf(key);
  if (num != -1){
    res_val = num + 1;
  }
  if (res_val > key_arr.length){
    return false;
  }
  return key_arr[res_val];
}

function requestApi(url, key, key_arr){
  var res_val;
  const QUOTA_EXCEEDED_1 = "Daily Limit Exceeded.";
  const QUOTA_EXCEEDED_2 = 'quotaExceeded';
  const QUOTA_EXCEEDED_3 = "Daily Limit for Unauthenticated Use Exceeded. Continued use requires signup.";
  
  var q_suc = false;
  
  do{
    q_suc = false;
    
    try{
      res_val = UrlFetchApp.fetch(url + '&key=' + key.pri);//,{ muteHttpExceptions:true });
    }catch(e){
      Browser.msgBox('エラー：' + e.message);
      if(e.message.indexOf(QUOTA_EXCEEDED_1) == -1 && e.message.indexOf(QUOTA_EXCEEDED_2) == -1){
        Browser.msgBox('エラー：APIキーか、URLが無効です。');
        return false;
      }else{
        key.pri = getNextKey(key_arr, key.pri);
        if(key.pri == false){
          Browser.msgBox('エラー：クオータが上限です。');
          return false;
        }else{
          q_suc = true;
          //Browser.msgBox('切り替えた');
        }
      }
    }
  }while(q_suc)
  return res_val;
}
