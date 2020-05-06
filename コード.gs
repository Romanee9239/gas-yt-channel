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
 *チャンネルURLから動画データを取得するマン
 *
 **********************************************/
function getMovieData () {
  
  var book = SpreadsheetApp.getActiveSpreadsheet();
  var sheetData = book.getSheetByName("基本シート");
  
  const APIURL_GET_CHANNEL  = 'https://www.googleapis.com/youtube/v3/channels?';
  const APIURL_GET_PLAYLIST = 'https://www.googleapis.com/youtube/v3/playlistItems?';
  const APIURL_GET_VIDEODATA = 'https://www.googleapis.com/youtube/v3/videos?';
  const CHANNEL_URL_HEAD_DF = 'https://www.youtube.com/channel/';
  const CHANNEL_URL_HEAD_US = 'https://www.youtube.com/user/';


  var inputCelCol = 2;
  var inputCelRow = 2;
  
  var colURL = 6;
  var colContributeCount = 12;
  var colIcon = 4;

  var rowStartData = 6
  var rowEndData = sheetData.getDataRange().getLastRow()
  

  //設定値取得
  var c = 5;
  var totalViewBuzzCond = sheetData.getRange(c++, 2).getValue();  //バズったと思う視聴回数
  var resentPeriodCond   = sheetData.getRange(c++, 2).getValue();  //初期バズ観察日数
  var periodBuzzCond    = sheetData.getRange(c++, 2).getValue();  //初期バズ判定視聴数
  
  //apiキー取得
  var key = {pri : ''};
  var key_arr = [];
  
  var lastRow = sheetData.getRange(sheetData.getMaxRows(), 2).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  var counta = 0;
  for(var i = 10; i <= lastRow; i++){
    key_arr[counta] = sheetData.getRange(i, 2).getValue();
    if(key_arr[counta] != ''){ 
      counta++; 
    }
  }
  key.pri = key_arr[0];

  //チャンネルID取得
  var isUser = false;
  var cnlUrl = sheetData.getRange(2, 1).getValue();
  var cid;
  if(cnlUrl.indexOf(CHANNEL_URL_HEAD_DF) == 0){
    cid = cnlUrl.replace(CHANNEL_URL_HEAD_DF,'');
  }else if(cnlUrl.indexOf(CHANNEL_URL_HEAD_US) == 0){
    cid = cnlUrl.replace(CHANNEL_URL_HEAD_US,'');
  }else{
    Browser.msgBox('エラー：チャンネルのURLじゃありません');
    return;
  }
  
  var cid_arr = cid.split('/');
  
  var id = cid_arr[0];
  
  var prm = "part=contentDetails,snippet,statistics"
  + (isUser ? "&forName=" : "&id=") + id;
  
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
  
  var dataURL = APIURL_GET_CHANNEL + prm;
  
  var resp;
  resp = requestApi(dataURL, key, key_arr);
  if (resp == false){ return; }
  var c_json  = JSON.parse(resp.getContentText());
  var c_items = c_json.items[0];
  var c_snippet = c_items.snippet;
  var c_statistics = c_items.statistics;
  var c_contentDetails = c_items.contentDetails;
  
  //チャンネル名
  ctitle = c_snippet.title;
  //登録者数
  var subs;
  if(c_statistics.hiddenSubscriberCount){
    //非公開
    subs = '非公開';
  }else{
    subs = c_statistics.subscriberCount;
  }
  //総再生数
  var c_view = c_statistics.viewCount;
  
  //動画一覧取得
  prm = "part=snippet&maxResults=50&playlistId=" + c_contentDetails.relatedPlaylists.uploads;
  
  var playlistUrl = APIURL_GET_PLAYLIST + prm;
  var res = requestApi(playlistUrl, key, key_arr);
  if (res == false){ return; }
  var jp = JSON.parse(res.getContentText());
  
  var next = jp.nextPageToken;
  var total = jp.pageInfo.totalResults;
  var list = jp.items;
  
  //動画情報取得
  var count = 0;
  
  //実行日の日付
  var todate = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd");
  
end_label:
  for(var j = 0; j<(total/50); j+=1){
    
    if (j>0){
      var nextUrl = playlistUrl + "&pageToken=" + next;
      res = requestApi(nextUrl, key, key_arr);
      if (res == false){ return; }
      jp = JSON.parse(res.getContentText());
      list = jp.items;
      next = jp.nextPageToken;
    }
    for(var k = count; k<count + list.length; k+=1){
      //ID
      targetVideoId[k] = list[k - count].snippet.resourceId.videoId;
      
      var vprm = "&part=snippet,contentDetails,statistics&id=" + targetVideoId[k];
      
      vjson = requestApi(APIURL_GET_VIDEODATA + vprm, key, key_arr);
      if (vjson == false){ return; }
      var data = JSON.parse(vjson.getContentText()).items[0];
      
      //url
      targetVideoUrl[k] = "https://www.youtube.com/watch?v=" + targetVideoId[k];
      //タイトル
      targetVideoTitle[k] = data.snippet.title;
      //日時
      var publishedTime = data.snippet.publishedAt;
      targetVideoDate[k] = Utilities.formatDate(new Date(publishedTime), "JST", "yyyy/MM/dd");
      targetVideoTime[k] = Utilities.formatDate(new Date(publishedTime), "JST", "HH:mm:ss");
      
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
    }
    count = k;
  }

  
  //チャンネルのシート取得
  var newSheet = book.getSheetByName(ctitle);
  if(newSheet != null){
    book.deleteSheet(newSheet);
  }
  book.insertSheet(ctitle);
  newSheet = book.getSheetByName(ctitle);
  
  //チャンネル名
  var o = 1;
  newSheet.getRange(o++, 1).setValue(ctitle);
  newSheet.getRange(o++, 1).setValue('チャンネル登録者数：' + subs);
  newSheet.getRange(o++, 1).setValue('総再生数：' + c_view);
  
  var m = 1;
  
  newSheet.getRange(o, m++).setValue('タイトル');
  var date = m;
  newSheet.getRange(o, m++).setValue('投稿日');
  newSheet.getRange(o, m++).setValue('曜日');
  newSheet.getRange(o, m++).setValue('投稿時間');
  var view = m;  
  newSheet.getRange(o, m++).setValue('再生数');
  newSheet.getRange(o, m++).setValue('日次再生');
  var h = m;
  newSheet.getRange(o, m++).setValue('動画時間');
  newSheet.getRange(o, m++).setValue('高評価数');
  newSheet.getRange(o, m++).setValue('低評価数');
  newSheet.getRange(o, m++).setValue('コメント');
  var tags = m;
  newSheet.getRange(o, m++).setValue('タグ');
  var tmp = m;
  
  var tempCel = newSheet.getRange(1, 2);
  tempCel.setValue(todate);
  
  //二次元配列の生成
  var targetValue = [];
  var ll = 0;
  for(var l = 0; l < targetVideoId.length; l+=1){
    if (targetVideoTags[l] == null){
      //none
    }else{
      var tag = targetVideoTags[l][0];
      for(var t=1; t<targetVideoTags[l].length; t+=1){
        tag = tag + ',' + targetVideoTags[l][t];
      }
    }
    Logger.log(targetVideoTitle[l]);
    
    ll = l + o + 1;
    
    targetValue[l] = [
      '=HYPERLINK("' + targetVideoUrl[l] + '","' + targetVideoTitle[l].split('"').join('""') + '")',
      targetVideoDate[l],
      '=text(B' + (ll) + ',"ddd")',
      targetVideoTime[l],
      targetVideoView[l],
      '=E' + (ll) + '/(today() + 1 - B' + (ll) + ')',
      targetVideoDuration[l],
      targetLikeCount[l],
      targetDislikeCount[l],
      targetComCount[l],
      tag,
      '=B' + (ll) + '-$B$1'
    ];
  }
  
  newSheet.getRange(o + 1,1,targetValue.length,targetValue[0].length).setValues(targetValue);
  
  for(var n = 0; n<targetValue.length; n+=1){  
  
    var nn = n + o + 1;
    
    var videoView = newSheet.getRange(nn, view);
    var videoDate = newSheet.getRange(nn, date);
    
    //最近出てすぐ伸びた動画は日付をハイライトする
    var tempDate = newSheet.getRange(nn, tmp).getValue();
    if (tempDate * -1 <= resentPeriodCond && videoView.getValue() > periodBuzzCond){
      videoDate.setBackground('#ffff00');
      videoDate.setFontStyle('Bold');
    }
    //特定回数以上再生された動画は再生回数をハイライトする
    if (videoView.getValue() > totalViewBuzzCond){
      videoView.setBackground('#ffff00');
      videoView.setFontStyle('Bold');
    }
    
  }  
  
  newSheet.getRange(4, tmp, nn, 1).clear();
  
  newSheet.setFrozenRows(o);
  
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
