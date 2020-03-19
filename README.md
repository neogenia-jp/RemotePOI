# RemotePOI
The gRPC Service for IO to xls/xlsx files using NPOI.


How to run ruby sample client.

```
gem install grpc grpc-tools
cd ruby_client
ruby ./rpoi_client.rb '192.168.1.158:37722'
```

to re-generate grpc defines.

```
grpc_tools_ruby_protoc -I../XlsManiSvc/XlsManiSvc/Protos --ruby_out=lib --grpc_out=lib ../XlsManiSvc/XlsManiSvc/Protos/rpoi.proto
```

