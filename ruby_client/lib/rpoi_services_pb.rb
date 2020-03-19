# Generated by the protocol buffer compiler.  DO NOT EDIT!
# Source: rpoi.proto for package 'rpoi'

require 'grpc'
require_relative 'rpoi_pb'

module Rpoi
  module RemotePOI
    # The remote POI service definition.
    class Service

      include GRPC::GenericService

      self.marshal_class_method = :encode
      self.unmarshal_class_method = :decode
      self.service_name = 'rpoi.RemotePOI'

      # Load & Save the Book.
      rpc :UploadTemplate, Google::Protobuf::BytesValue, Google::Protobuf::Empty
      rpc :UseTemplateFile, Google::Protobuf::StringValue, Google::Protobuf::Empty
      rpc :Download, Google::Protobuf::Empty, Google::Protobuf::BytesValue
      # Control the sheet.
      rpc :GetSheetsCount, Google::Protobuf::Empty, Google::Protobuf::Int32Value
      rpc :GetSheetName, Google::Protobuf::Int32Value, Google::Protobuf::StringValue
      rpc :SelectSheetAt, Google::Protobuf::Int32Value, Google::Protobuf::Empty
      rpc :CreateSheet, Google::Protobuf::StringValue, Google::Protobuf::Empty
      # Cell value / type
      rpc :GetCellValueType, CellAddress, CellValueType
      rpc :GetCellValue, CellAddress, CellValue
      rpc :SetCellValue, CellAddressWithValue, Google::Protobuf::Empty
    end

    Stub = Service.rpc_stub_class
  end
end
