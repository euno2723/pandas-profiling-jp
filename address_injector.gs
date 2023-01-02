function mergePrint() {

  //シート取得
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();

  //プレゼンテーション指定
  //https://docs.google.com/presentation/d/スライドのID/edit#
  var presentation = SlidesApp.openById('スライドのID');


  var cnt_slide = 0;

  //シート2行目から1行ずつ繰り返し
  for(var i=1; i<data.length; i++){


    //喪中等の理由で年賀状を送らない人はidを空欄にするとスキップできる
    if(!data[i][0]){
      ;
    } else if(data[i][0]){

      //スライドを1枚追加＆追加したページを取得
      presentation.appendSlide();  
      var slide = presentation.getSlides()[cnt_slide];
      cnt_slide = cnt_slide + 1

      //テキストボックス設置　insertShape(shapeType, left, top, width, height)
      var zip = [];
      zip[0] = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 121,  26,  20,  30); //郵便番号
      zip[1] = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 141,  26,  20,  30); //郵便番号
      zip[2] = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 161,  26,  20,  30); //郵便番号
      zip[3] = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 182,  26,  20,  30); //郵便番号
      zip[4] = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 202,  26,  20,  30); //郵便番号
      zip[5] = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 221,  26,  20,  30); //郵便番号
      zip[6] = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 241,  26,  20,  30); //郵便番号
      var addres_1 = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX,  25, 150, 240,  20); //住所１
      var addres_2 = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX,  25, 170, 240,  20); //住所２
      var name_1 = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX,   0, 200, 283,  40); //お名前
      var name_2 = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX,   0, 230, 283,  40); //連名

      //テキストボックスに入力
      //郵便番号
      for(var j=0; j<7; j++){
        zip[j].getText().setText(String(data[i][1]).charAt(j)).getTextStyle().setFontFamily("Sawarabi Mincho").setFontSize(20);
      }

      //住所１
      addres_1.getText().setText(data[i][2]).getTextStyle().setFontFamily("Sawarabi Mincho").setFontSize(13);

      //住所２（空欄なら飛ばす）
      if(data[i][3]){
        addres_2.getText().setText(data[i][3]).getTextStyle().setFontFamily("Sawarabi Mincho").setFontSize(13);
      }

      //お名前（太字＆中央揃え）
      name_1.getText().setText(data[i][4] + ' ' + data[i][5] + ' ' + data[i][6]).getTextStyle().setFontFamily("Sawarabi Mincho").setFontSize(25);
      name_1.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

      //連名（太字＆中央揃え）
      if(data[i][5]){
        var space_length = data[i][4].length
        var space_zenkaku = '　'
        var space = space_zenkaku.repeat(space_length);

        name_2.getText().setText(space + ' ' + data[i][7]+ ' ' + data[i][8]).getTextStyle().setFontFamily("Sawarabi Mincho").setFontSize(25);
        name_2.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
      }
    }
  }   
}







