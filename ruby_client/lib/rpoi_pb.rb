# frozen_string_literal: true
# Generated by the protocol buffer compiler.  DO NOT EDIT!
# source: rpoi.proto

require 'google/protobuf'

require 'google/protobuf/empty_pb'
require 'google/protobuf/timestamp_pb'
require 'google/protobuf/wrappers_pb'

descriptor_data = "\n\nrpoi.proto\x12\x04rpoi\x1a\x1bgoogle/protobuf/empty.proto\x1a\x1fgoogle/protobuf/timestamp.proto\x1a\x1egoogle/protobuf/wrappers.proto\"4\n\rCellValueType\x12#\n\x05value\x18\x01 \x01(\x0e\x32\x14.rpoi.CellValueTypes\"\xd1\x01\n\tCellValue\x12\'\n\tvalueType\x18\x01 \x01(\x0e\x32\x14.rpoi.CellValueTypes\x12\x10\n\x08is_blank\x18\x02 \x01(\x08\x12\x15\n\rnumeric_value\x18\x03 \x01(\x01\x12\x14\n\x0cstring_value\x18\x04 \x01(\t\x12\x12\n\nbool_value\x18\x05 \x01(\x08\x12\x33\n\x0f\x64\x61te_time_value\x18\x06 \x01(\x0b\x32\x1a.google.protobuf.Timestamp\x12\x13\n\x0b\x65rror_value\x18\x07 \x01(\r\"\'\n\x0b\x43\x65llAddress\x12\x0b\n\x03row\x18\x01 \x01(\x05\x12\x0b\n\x03\x63ol\x18\x02 \x01(\x05\"P\n\x14\x43\x65llAddressWithValue\x12\x0b\n\x03row\x18\x01 \x01(\x05\x12\x0b\n\x03\x63ol\x18\x02 \x01(\x05\x12\x1e\n\x05value\x18\x03 \x01(\x0b\x32\x0f.rpoi.CellValue\"+\n\x0cIndexAndName\x12\r\n\x05index\x18\x01 \x01(\x05\x12\x0c\n\x04name\x18\x02 \x01(\t\"+\n\x0cIndexAndFlag\x12\r\n\x05index\x18\x01 \x01(\x05\x12\x0c\n\x04\x66lag\x18\x02 \x01(\x08\"?\n\rIndexAndState\x12\r\n\x05index\x18\x01 \x01(\x05\x12\x1f\n\x05state\x18\x02 \x01(\x0e\x32\x10.rpoi.SheetState\"A\n\x13RecalculationPolicy\x12*\n\x05value\x18\x01 \x01(\x0e\x32\x1b.rpoi.RecalculationPolicies*}\n\x0e\x43\x65llValueTypes\x12\x0b\n\x07Numeric\x10\x00\x12\n\n\x06String\x10\x01\x12\x0b\n\x07\x46ormula\x10\x02\x12\t\n\x05\x42lank\x10\x03\x12\x0b\n\x07\x42oolean\x10\x04\x12\x0c\n\x08\x44\x61teTime\x10\x05\x12\t\n\x05\x45rror\x10\x06\x12\x14\n\x07Unknown\x10\xff\xff\xff\xff\xff\xff\xff\xff\xff\x01*5\n\nSheetState\x12\x0b\n\x07Visible\x10\x00\x12\n\n\x06Hidden\x10\x01\x12\x0e\n\nVeryHidden\x10\x02*A\n\x15RecalculationPolicies\x12\x08\n\x04None\x10\x00\x12\x0b\n\x07SetFlag\x10\x01\x12\x11\n\rForceEvaluate\x10\x02\x32\x91\n\n\tRemotePOI\x12\x45\n\x0eUploadTemplate\x12\x1b.google.protobuf.BytesValue\x1a\x16.google.protobuf.Empty\x12G\n\x0fUseTemplateFile\x12\x1c.google.protobuf.StringValue\x1a\x16.google.protobuf.Empty\x12?\n\x08\x44ownload\x12\x16.google.protobuf.Empty\x1a\x1b.google.protobuf.BytesValue\x12K\n\x16GetRecalculationPolicy\x12\x16.google.protobuf.Empty\x1a\x19.rpoi.RecalculationPolicy\x12K\n\x16SetRecalculationPolicy\x12\x19.rpoi.RecalculationPolicy\x1a\x16.google.protobuf.Empty\x12\x45\n\x0eGetSheetsCount\x12\x16.google.protobuf.Empty\x1a\x1b.google.protobuf.Int32Value\x12I\n\x0cGetSheetName\x12\x1b.google.protobuf.Int32Value\x1a\x1c.google.protobuf.StringValue\x12\x44\n\rSelectSheetAt\x12\x1b.google.protobuf.Int32Value\x1a\x16.google.protobuf.Empty\x12\x43\n\x0b\x43reateSheet\x12\x1c.google.protobuf.StringValue\x1a\x16.google.protobuf.Empty\x12\x44\n\rRemoveSheetAt\x12\x1b.google.protobuf.Int32Value\x1a\x16.google.protobuf.Empty\x12:\n\x0cSetSheetName\x12\x12.rpoi.IndexAndName\x1a\x16.google.protobuf.Empty\x12=\n\x0eSetSheetHidden\x12\x13.rpoi.IndexAndState\x1a\x16.google.protobuf.Empty\x12=\n\nCloneSheet\x12\x12.rpoi.IndexAndName\x1a\x1b.google.protobuf.Int32Value\x12\x41\n\nClearRowAt\x12\x1b.google.protobuf.Int32Value\x1a\x16.google.protobuf.Empty\x12@\n\x12SetRowZeroHeightAt\x12\x12.rpoi.IndexAndFlag\x1a\x16.google.protobuf.Empty\x12:\n\x10GetCellValueType\x12\x11.rpoi.CellAddress\x1a\x13.rpoi.CellValueType\x12\x32\n\x0cGetCellValue\x12\x11.rpoi.CellAddress\x1a\x0f.rpoi.CellValue\x12\x42\n\x0cSetCellValue\x12\x1a.rpoi.CellAddressWithValue\x1a\x16.google.protobuf.Empty\x12\x43\n\x0cGetMaxRowNum\x12\x16.google.protobuf.Empty\x1a\x1b.google.protobuf.Int32ValueB\r\xaa\x02\nXlsManiSvcb\x06proto3"

pool = Google::Protobuf::DescriptorPool.generated_pool

begin
  pool.add_serialized_file(descriptor_data)
rescue TypeError
  # Compatibility code: will be removed in the next major version.
  require 'google/protobuf/descriptor_pb'
  parsed = Google::Protobuf::FileDescriptorProto.decode(descriptor_data)
  parsed.clear_dependency
  serialized = parsed.class.encode(parsed)
  file = pool.add_serialized_file(serialized)
  warn "Warning: Protobuf detected an import path issue while loading generated file #{__FILE__}"
  imports = [
    ["google.protobuf.Timestamp", "google/protobuf/timestamp.proto"],
  ]
  imports.each do |type_name, expected_filename|
    import_file = pool.lookup(type_name).file_descriptor
    if import_file.name != expected_filename
      warn "- #{file.name} imports #{expected_filename}, but that import was loaded as #{import_file.name}"
    end
  end
  warn "Each proto file must use a consistent fully-qualified name."
  warn "This will become an error in the next major version."
end

module Rpoi
  CellValueType = ::Google::Protobuf::DescriptorPool.generated_pool.lookup("rpoi.CellValueType").msgclass
  CellValue = ::Google::Protobuf::DescriptorPool.generated_pool.lookup("rpoi.CellValue").msgclass
  CellAddress = ::Google::Protobuf::DescriptorPool.generated_pool.lookup("rpoi.CellAddress").msgclass
  CellAddressWithValue = ::Google::Protobuf::DescriptorPool.generated_pool.lookup("rpoi.CellAddressWithValue").msgclass
  IndexAndName = ::Google::Protobuf::DescriptorPool.generated_pool.lookup("rpoi.IndexAndName").msgclass
  IndexAndFlag = ::Google::Protobuf::DescriptorPool.generated_pool.lookup("rpoi.IndexAndFlag").msgclass
  IndexAndState = ::Google::Protobuf::DescriptorPool.generated_pool.lookup("rpoi.IndexAndState").msgclass
  RecalculationPolicy = ::Google::Protobuf::DescriptorPool.generated_pool.lookup("rpoi.RecalculationPolicy").msgclass
  CellValueTypes = ::Google::Protobuf::DescriptorPool.generated_pool.lookup("rpoi.CellValueTypes").enummodule
  SheetState = ::Google::Protobuf::DescriptorPool.generated_pool.lookup("rpoi.SheetState").enummodule
  RecalculationPolicies = ::Google::Protobuf::DescriptorPool.generated_pool.lookup("rpoi.RecalculationPolicies").enummodule
end
