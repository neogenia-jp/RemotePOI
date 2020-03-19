# RemotePOI
The gRPC Service for IO to xls/xlsx files using NPOI.

## ASP.NET Core gRPC service

How to build service as docker image.

```
docker build --no-cache -t rpoi_svc .
docker run -ti -p 37722:80 -v ${PWD}/sample_template:/mnt rpoi_svc
```

## Ruby client

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

