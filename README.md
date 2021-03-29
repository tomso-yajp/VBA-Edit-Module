read me

vbproject.vbcomponents.codemodule を操作します

module:abc_sub <br/>
（対象モジュール：abc_key）<br/>

setkey(var)：<br/>
変数の値を変更します <br/>
setkey("変数名,変数の型,変数の値") <br/>
※var の引数は、ダブルクォーテーションで囲みます <br/>

rowcount：<br/>
文字ジュール内の最終行を返します <br/>

checkkey(var)：<br/>
指定した変数の有無を確認します <br/>
checkkey("変数名") <br/>

getkey(var,n)：<br/>
指定した変数の行、または、値を返します <br/>
getkey("変数名",数値) <br/>
※第2引数：0、行を返します <br/>
※第2引数：1、値を返します <br/>

delkey(var)：<br/>
指定した変数を削除します <br/>
delkey("変数名") <br/>

dellines：<br/>
全ての行を削除します <br/>





