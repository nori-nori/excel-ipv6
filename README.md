#ExcelでIPv6


## What' this?

* Excel上でIPv6表記の正誤を判定したい
* 表記を統一したい（RFC5952?）

という二つの目的でinet_ntop4/6、inet_pton4/6をVBAにポーティングをしたもの



## Usage

ポーティングしたものは以下の４つ
  
	Boolean inet_ntop4(Integer(3), String)
	Boolean inet_ntop6(Integer(15), String)
	Boolean inet_pton4(String, Integer(3))
	Boolean inet_pton6(String, Integer(15))


それと直接呼べるように二つのユーザ定義関数を作成

	Boolean v6validate(String)
	Boolean v6normalize(String)





## Sample

### Sheet1 ユーザ定義関数版


Ａ列にアドレスを入力すると

* Ｂ列にIPv6アドレスの真偽（TRUE/FALSE）
* Ｃ列にRFC5952表記（FALSEの場合は"invalid"）
* IPv6アドレスでなければ入力セルの背景色を変更（条件付き書式）

を設定する



### Sheet2 イベントプロシージャ版

Ａ列にアドレスを入力すると

* Ｂ列にIPv6アドレスの真偽（TRUE/FALSE）
* IPv6アドレスでなければ入力セルの背景色を変更

を設定する



## その他
ライセンスはBSDあたりで

ほかに便利な関数があればコード付きで送って下さい:-)

LLに慣れすぎるとVBAはちかれます。。。


