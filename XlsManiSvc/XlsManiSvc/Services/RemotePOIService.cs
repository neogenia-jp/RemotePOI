using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Caching.Memory;
using Grpc.Core;
using Google.Protobuf;
using Google.Protobuf.WellKnownTypes;
using Endless;

namespace XlsManiSvc
{
    public class RemotePOIService : RemotePOI.RemotePOIBase
    {
        private readonly ILogger<RemotePOIService> _logger;
        private readonly IMemoryCache _cache;

        public RemotePOIService(ILogger<RemotePOIService> logger, IMemoryCache cache)
        {
            _logger = logger;
            _cache = cache;
        }

        private NPoiWrapper GetOrCreateWrapper(ServerCallContext context)
        {
            var key = context.Peer;
            return _cache.GetOrCreate(key, entry =>
            {
                _logger.LogInformation("New connection comming! {0}", context.Peer);
                entry.AbsoluteExpirationRelativeToNow = TimeSpan.FromSeconds(10);
                entry.PostEvictionCallbacks.Add(new PostEvictionCallbackRegistration { EvictionCallback = (k,v,r,s) => _DisposeCache(v as NPoiWrapper) });
                return new NPoiWrapper(_logger);
            });
        }

        private void _DisposeCache(NPoiWrapper w)
        {
            _logger.LogTrace("Disposing old npoi object. {0} created {1}. Current memory usage: {2} bytes", w, w.CreatedAt, Environment.WorkingSet);
            if (w == null) return;
            w.Dispose();
            GC.Collect();
        }

        public override Task<Empty> UploadTemplate(BytesValue data, ServerCallContext context)
            => Task.Factory.StartNew(() =>
            {
                _logger.LogDebug("UploadTemplate() called with {0} bytes data.", data.Value.Length);
                GetOrCreateWrapper(context).LoadTemplateFromData(data.Value.ToStream());
                return new Empty();
            });

        public override Task<Empty> UseTemplateFile(StringValue path, ServerCallContext context)
            => Task.Factory.StartNew(() =>
            {
                _logger.LogDebug("UseTemplateFile({0})", path);
                GetOrCreateWrapper(context).LoadTemplateFromFile(path.Value);
                return new Empty();
            });

        public override Task<BytesValue> Download(Empty _, ServerCallContext context)
            => Task.Factory.StartNew(() =>
            {
                _logger.LogDebug("Download()");
                var ret = new BytesValue { Value = ByteString.FromStream(GetOrCreateWrapper(context).Download()) };
                _logger.LogTrace("  --------------------- Current memory usage: {0} bytes ---------------------", Environment.WorkingSet);
                return ret;
            });

        public override Task<Int32Value> GetSheetsCount(Empty _, ServerCallContext context)
            => Task.Factory.StartNew(() =>
            {
                _logger.LogDebug("GetSheetsCount()");
                return new Int32Value { Value = GetOrCreateWrapper(context).GetSheetCount() };
            });

        public override Task<StringValue> GetSheetName(Int32Value arg, ServerCallContext context)
            => Task.Factory.StartNew(() =>
            {
                _logger.LogDebug("GetSheetName({0})", arg);
                return new StringValue { Value = GetOrCreateWrapper(context).GetSheetName(arg.Value) };
            });

        public override Task<Empty> SelectSheetAt(Int32Value arg, ServerCallContext context)
            => Task.Factory.StartNew(() =>
            {
                _logger.LogDebug("SelectSheetAt({0})", arg);
                GetOrCreateWrapper(context).SelectSheetAt(arg.Value);
                return new Empty();
            });

        public override Task<Empty> CreateSheet(StringValue name, ServerCallContext context)
            => Task.Factory.StartNew(() =>
            {
                _logger.LogDebug("CreateSheet({0})", name);
                GetOrCreateWrapper(context).CreateSheet(name.Value);
                return new Empty();
            });

        public override Task<CellValueType> GetCellValueType(CellAddress addr, ServerCallContext context)
            => Task.Factory.StartNew(() =>
            {
                _logger.LogDebug("GetCellValueType(row:{0}, col:{1})", addr.Row, addr.Col);
                return new CellValueType { Value = GetOrCreateWrapper(context).GetCellValueType(addr) };
            });

        public override Task<CellValue> GetCellValue(CellAddress addr, ServerCallContext context)
            => Task.Factory.StartNew(() =>
            {
                _logger.LogDebug("GetCellValue(row:{0}, col:{1})", addr.Row, addr.Col);
                return GetOrCreateWrapper(context).GetCellValue(addr);
            });

        public override Task<Empty> SetCellValue(CellAddressWithValue addrv, ServerCallContext context)
            => Task.Factory.StartNew(() =>
            {
                _logger.LogDebug("SetCellValue(row:{0}, col:{1})", addrv.Row, addrv.Col, addrv.Value.Inspect());
                GetOrCreateWrapper(context).SetCellValue(addrv);
                return new Empty();
            });
    }
}
