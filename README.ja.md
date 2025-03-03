# RemotePOI

RemotePOI は NPOI を使った xls/xlsx ファイルの操作を提供する gRPCサービスです。
[![build and test](https://github.com/neogenia-jp/RemotePOI/actions/workflows/build-and-test.yml/badge.svg)](https://github.com/neogenia-jp/RemotePOI/actions/workflows/build-and-test.yml)

## Quick start

### サービスの起動

Dockerコンテナを使ってサービスを起動できます。

```bash
cd /path/to/this/repository
docker build -t rpoi_svc ./XlsManiSvc/
docker run -ti -p 37722:80 -v ${PWD}/sample_template:/mnt rpoi_svc
```

これでコンテナが起動し、ポート番号 37722 で gRPCサービスが待ち受けしています。
ASP.NET Core gRPC サービスを利用してホスティングしています。


### サンプルクライアントの実行

`ruby_client/` ディレクトリには Ruby で書かれたクライアントのサンプルコードが置いてあります。

サンプルは、`sample_template/est_template.xls` ファイルを開き、いくつかのセルに値をセットして
保存する、といった内容です。
以下のようにして実行できます（192.168.1.158 の部分はサービスが稼働しているホストのアドレスに読み替えてください）。

```bash
gem install grpc grpc-tools
cd ruby_client/
ruby ./rpoi_client.rb '192.168.1.158:37722' /tmp/sample.xls
```

これで `/tmp/sample.xls` にファイルが保存されているはずです。


## マルチセッションモード

サービス起動時に環境変数 `ENABLE_MULTI_SESSION` を指定することでマルチセッションモードが有効になります。
マルチセッションモードでない場合は、１つの接続元から複数のセッションでアクセスしようした場合に、
サービスはセッションを区別できなくなり、正しく動作しません。
並列アクセスが想定されるケースでは必ずマルチセッションモードで使用してください。

クライアント側では、初回アクセス時にサービスより `x-session-id` というHTTPヘッダが付与されるので、
２回目以降のアクセスにはそのヘッダをそのまま付けてリクエストする必要があります。

上記のサンプルが `ruby_client/rpoi_client2.rb` に入っています。
同一のアクセス元（IPアドレス、ポート番号）から２つのセッションを同時に開き、並列的にファイル操作を行うような
内容になっています。

以下のようにして実行できます。

サービスの起動
```bash
docker build -t rpoi_svc ./XlsManiSvc/
docker run -ti -p 37722:80 -v ${PWD}/sample_template:/mnt -e ENABLE_MULTI_SESSION=1 rpoi_svc
```

クライアントの起動
```bash
ruby ./rpoi_client2.rb '192.168.1.158:37722' /tmp/sample.xls
```

## gRPCの定義を再生成する方法

.NET の方はビルド時に自動的に再生成されます。

Ruby クライアントの方は、以下のコマンドで再生成できます。

```bash
grpc_tools_ruby_protoc -I../XlsManiSvc/XlsManiSvc/Protos --ruby_out=lib --grpc_out=lib ../XlsManiSvc/XlsManiSvc/Protos/rpoi.proto
```

## 付録

gRPC でのHTTPヘッダの取り扱いについて役に立ったリンク集

- https://www.wantedly.com/companies/wantedly/post_articles/219429
- https://tech.raksul.com/2019/06/05/grpc-client-interceptor-intro-with-ruby/

