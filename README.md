# RemotePOI

RemotePOI is a gRPC service that provides for the manipulation of xls/xlsx files using NPOI.
[![build and test](https://github.com/neogenia-jp/RemotePOI/actions/workflows/build-and-test.yml/badge.svg)](https://github.com/neogenia-jp/RemotePOI/actions/workflows/build-and-test.yml)

[日本語版](./README.ja.md)

## Quick start

### Starting the Service.

You can use a Docker container to start a service.

```bash
cd /path/to/this/repository
docker build -t rpoi_svc ./XlsManiSvc/
docker run -ti -p 37722:80 -v ${PWD}/sample_template:/mnt rpoi_svc
```

The container is now up and the gRPC service is waiting on port number 37722.
We are hosting it using the ASP.NET Core gRPC service.

### Running the Sample Client.

The `ruby_client/` directory contains sample code for a client written in Ruby.


The sample is to open a file `sample_template/est_template.xls`,
 set values in some cells and save them, and so on.
You can do this in the following way
 (Replace 192.168.1.158 with host's address of the service is running).

```bash
gem install grpc grpc-tools
cd ruby_client/
ruby ./rpoi_client.rb '192.168.1.158:37722' /tmp/sample.xls
```

Now the file should be saved to `/tmp/sample.xls`.

## Multi-session mode.

Multi-session mode is enabled by setting the environment variable `ENABLE_MULTI_SESSION`
when the service is started.
When not in multi-session mode, the service will not be able to distinguish
between multiple sessions when a user tries to access the service from a single source.
You should always use multi-session mode if you expect parallel access to the service.

On the client side, the service gives you an HTTP header `x-session-id`
when you access the service for the first time,
so you need to add it to your request on the second and subsequent accesses.

The above sample is in `ruby_client/rpoi_client2.rb`,
and you can find it in The following is an example.
It is designed to open two sessions simultaneously from the same access source
(IP address and port number) to perform file operations in parallel.

You can do this as follows:

Starting the service
```bash
docker build -t rpoi_svc ./XlsManiSvc/
docker run -ti -p 37722:80 -v ${PWD}/sample_template:/mnt -e ENABLE_MULTI_SESSION=1 rpoi_svc
```

Run the sample client
```bash
ruby ./rpoi_client2.rb '192.168.1.158:37722' /tmp/sample.xls
```

## How to regenerate the gRPC definition

NET, it is automatically regenerated at build time.

For Ruby clients, you can regenerate it with the following commands

```bash
grpc_tools_ruby_protoc -I../XlsManiSvc/XlsManiSvc/Protos --ruby_out=lib --grpc_out=lib ../XlsManiSvc/XlsManiSvc/Protos/rpoi.proto
```

## Appendix.

Useful links on handling HTTP headers in gRPC.
- https://www.wantedly.com/companies/wantedly/post_articles/219429
- https://tech.raksul.com/2019/06/05/grpc-client-interceptor-intro-with-ruby/
