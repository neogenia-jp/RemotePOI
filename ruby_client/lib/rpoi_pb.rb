# Generated by the protocol buffer compiler.  DO NOT EDIT!
# source: rpoi.proto

require 'google/protobuf'

require 'google/protobuf/empty_pb'
require 'google/protobuf/timestamp_pb'
require 'google/protobuf/wrappers_pb'
Google::Protobuf::DescriptorPool.generated_pool.build do
  add_file("rpoi.proto", :syntax => :proto3) do
    add_message "rpoi.CellValueType" do
      optional :value, :enum, 1, "rpoi.CellValueTypes"
    end
    add_message "rpoi.CellValue" do
      optional :valueType, :enum, 1, "rpoi.CellValueTypes"
      optional :is_blank, :bool, 2
      optional :numeric_value, :double, 3
      optional :string_value, :string, 4
      optional :bool_value, :bool, 5
      optional :date_time_value, :message, 6, "google.protobuf.Timestamp"
      optional :error_value, :uint32, 7
    end
    add_message "rpoi.CellAddress" do
      optional :row, :int32, 1
      optional :col, :int32, 2
    end
    add_message "rpoi.CellAddressWithValue" do
      optional :row, :int32, 1
      optional :col, :int32, 2
      optional :value, :message, 3, "rpoi.CellValue"
    end
    add_message "rpoi.IndexAndName" do
      optional :index, :int32, 1
      optional :name, :string, 2
    end
    add_enum "rpoi.CellValueTypes" do
      value :Numeric, 0
      value :String, 1
      value :Formula, 2
      value :Blank, 3
      value :Boolean, 4
      value :DateTime, 5
      value :Error, 6
      value :Unknown, -1
    end
  end
end

module Rpoi
  CellValueType = ::Google::Protobuf::DescriptorPool.generated_pool.lookup("rpoi.CellValueType").msgclass
  CellValue = ::Google::Protobuf::DescriptorPool.generated_pool.lookup("rpoi.CellValue").msgclass
  CellAddress = ::Google::Protobuf::DescriptorPool.generated_pool.lookup("rpoi.CellAddress").msgclass
  CellAddressWithValue = ::Google::Protobuf::DescriptorPool.generated_pool.lookup("rpoi.CellAddressWithValue").msgclass
  IndexAndName = ::Google::Protobuf::DescriptorPool.generated_pool.lookup("rpoi.IndexAndName").msgclass
  CellValueTypes = ::Google::Protobuf::DescriptorPool.generated_pool.lookup("rpoi.CellValueTypes").enummodule
end
