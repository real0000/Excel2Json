# Excel2Json v0.1

相依列表如下:

1. xlsx i/o : https://brechtsanders.github.io/xlsxio/
2. freexl : https://www.gaia-gis.it/fossil/freexl/index
3. boost : http://www.boost.org/
4. rapidjson : https://github.com/Tencent/rapidjson

指令Excel2Json -f <設定檔> [-pretty]


設定檔請參照Example.xml 
# Tables部分事項：

1.底下必須列出會用到的欄位和表，一個表單可以包含複數資料表，每個資料表間必須在最上面一欄有欄位名稱，並指定在descRow

2.name則可以任意指定，但必須不重複

3.type有text、int、float三種

4.每張資料表之前必須隔開，連載一起會無法判斷
  
# Outputs部分事項：

1.若資料連結至其他表，table是必填欄位，且該表必須有跟連結欄位同名稱的資料(desc部分)

2.linkType可以為dict、array以及obj

3.linkType為dict時，需填key，以將該值作為字典索引
