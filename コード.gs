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

  var sheetData = book.getSheetByName("入力");

  var inputCelCol = 2;
  var inputCelRow = 2;
  
  var colURL = 6;
  var colContributeCount = 12;
  var colIcon = 4;

  var rowStartData = 6
  var rowEndData = sheetData.getDataRange().getLastRow()

  /*
  //動画一覧取得*/
  
  var cid = sheetData.getRange(2, 2).getValue().replace('https://www.youtube.com/channel/','');
  
  if(cid.indexOf('/') != -1){
    cid = cid.slice(cid.indexOf('/'));
  }
    
  	var url = "https://www.googleapis.com/youtube/v3/channels?"
	var key = "&key=AIzaSyC5yoDsNzrbRN8HIYMDYNEVg1g8RoZbEOo";
    var key = "&key=AIzaSyCrs4HN8huGIkFdh6Zt90WR6enJwlco_vY";
	var prm = "part=contentDetails,snippet"
			+ (cid ? "&id=" + cid : "");
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
    
    var dataURL = url + prm + key;
    var resp = UrlFetchApp.fetch(dataURL);//,{ muteHttpExceptions:true });
    //Logger.log(dataURL);
  
    //
        
        
    url = "https://www.googleapis.com/youtube/v3/playlistItems?";
    prm = "part=snippet&maxResults=50&playlistId=" + JSON.parse(resp.getContentText()).items[0].contentDetails.relatedPlaylists.uploads;
    
    dataURL = url + prm + key;
    res = UrlFetchApp.fetch(dataURL);//,{ muteHttpExceptions:true });
    Logger.log(dataURL);
    var jp = JSON.parse(res.getContentText());
    var next = jp.nextPageToken;
    Logger.log("72:" + next);
    var total = jp.pageInfo.totalResults;
    var list = jp.items;
  
    ctitle = JSON.parse(res.getContentText()).items[0].snippet.channelTitle;
    Logger.log(ctitle);
    
    //動画情報取得
  var count = 0;
  Logger.log("[jmax]:" + total/50);
  
  //実行日の日付
  var todate = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd");
  
  var keySheet = book.getSheetByName('key');
  //伸びたと判断する再生回数
  var totalViewBuzzCond = keySheet.getRange(2,12).getValue();
  //一週間に伸びたと判断する再生回数
  var oneWeekBuzzCond = keySheet.getRange(3,12).getValue();
  //マジックナンバーはよくないけど設定値増えてから対応する
  
end_label:
  for(var j = 0; j<(total/50); j+=1){
    
    if (j>0){
      var nextUrl = dataURL + "&pageToken=" + next;
      jp = JSON.parse(UrlFetchApp.fetch(nextUrl).getContentText());
      list = jp.items;
      next = jp.nextPageToken;
    }
    for(var k = count; k<count + list.length; k+=1){
      //ID
      targetVideoId[k] = list[k - count].snippet.resourceId.videoId;
      
      var vurl = "https://www.googleapis.com/youtube/v3/videos?"
      //var vprm = "&part=snippet,contentDetails,statistics,fileDetails&id=" + targetVideoId[k];
      var vprm = "&part=snippet,contentDetails,statistics&id=" + targetVideoId[k];
      
      vurl = vurl + vprm + key;
      
      try{
        vjson = UrlFetchApp.fetch(vurl);//,{ muteHttpExceptions:true });
        Logger.log(vurl);
      }catch(e){
        Browser.msgBox(k + '番目の動画でクォータが最大値になりました');
        Logger.log(e.message);
        k--;
        break end_label;
      }
      var data = JSON.parse(vjson.getContentText()).items[0];
      Logger.log(data);
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
      //ファイル名
      //targetFileName[k] = data.fileDetails.fileName
      //タグ取得
      targetVideoTags[k] = data.snippet.tags;
    }
    count = k;
  }    

  
  //チャンネルのシート取得
  var newSheet = book.getSheetByName(ctitle);
  if(newSheet == null){
    book.insertSheet(ctitle);
    newSheet = book.getSheetByName(ctitle);
  }
  var m = 1
  newSheet.getRange(1,1).setValue(ctitle);
  newSheet.getRange(2, m++).setValue('タイトル');
  var date = m;
  newSheet.getRange(2, m++).setValue('投稿日');
  newSheet.getRange(2, m++).setValue('曜日');
  newSheet.getRange(2, m++).setValue('投稿時間');
  var view = m;  
  newSheet.getRange(2, m++).setValue('再生数');
  newSheet.getRange(2, m++).setValue('日次再生');
  var h = m;
  newSheet.getRange(2, m++).setValue('動画時間');
  newSheet.getRange(2, m++).setValue('高評価数');
  newSheet.getRange(2, m++).setValue('低評価数');
  newSheet.getRange(2, m++).setValue('コメント');
  newSheet.getRange(2, m++).setValue('ファイル名');
  var tags = m;
  newSheet.getRange(2, m++).setValue('タグ');
  var tmp = m;
  
  var tempCel = newSheet.getRange(1, 3);
  tempCel.setValue(todate);
  
  //二次元配列の生成
  var targetValue = [];
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
    
    targetValue[l] = [
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
  }
  
  newSheet.getRange(3,1,targetValue.length,targetValue[0].length).setValues(targetValue);
  
  for(var n = 0; n<targetValue.length; n+=1){  
    
    var videoView = newSheet.getRange(n+3, view);
    var videoDate = newSheet.getRange(n+3, date);
    
    //一週間以内に伸びた動画は日付をハイライトする
    tempCel = newSheet.getRange(n+3, tmp);
    if (tempCel.getValue() > -8 && videoView.getValue() > oneWeekBuzzCond){
      //Logger.log('['+ l + ']:' + targetVideoDate[l].getDate() + ' - ' +  todate + ' = ' + thisDate - todate);
      videoDate.setBackground('#ffff00');
      videoDate.setFontStyle('Bold');
    }
    //特定回数以上再生された動画は再生回数をハイライトする
    if (videoView.getValue() > totalViewBuzzCond){
      videoView.setBackground('#ffff00');
      videoView.setFontStyle('Bold');
    }
  }  
  
  newSheet.getRange(3, tmp, n, 1).clear();
  
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
function getKey(keyCol){
  var res_val = [];
  var book = SpreadsheetApp.getActiveSpreadsheet();
  var sheetData = book.getSheetByName("key");
  
  var lastRow = sheetData.getRange(sheetData.getMaxRows(), keyCol).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  
  var res = "";
  var k = 0;
  
  for (var i = 2; i <= lastRow; i+=1){
    res = sheetData.getRange(i, keyCol).getValue();
    if (res != ""){
      res_val[k] = res;
      k++;
    }
  }
  
  return res_val;
}