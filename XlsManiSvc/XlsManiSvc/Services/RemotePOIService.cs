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

        private static int? _session_timeout_sec;
        private static bool? _enable_session;

        public RemotePOIService(ILogger<RemotePOIService> logger, IMemoryCache cache)
        {
            _logger = logger;
            _cache = cache;
            if (_enable_session == null)
            {
                switch (Environment.GetEnvironmentVariable("ENABLE_MULTI_SESSION"))
                {
                    case null:
                    case "0":
                    case "no":
                    case "false":
                    case "NO":
                    case "FALSE":
                        _enable_session = false;
                        break;
                    default:
                        _enable_session = true;
                        break;
                }
            }
            if (_session_timeout_sec == null)
            {
                _session_timeout_sec = int.Parse(Environment.GetEnvironmentVariable("SESSION_TIMEOUT_SEC") ?? "60");
                _logger.LogInformation("Multisession: {0}, Session timeout: {1} sec", _enable_session, _session_timeout_sec);
            }
        }

        private NPoiWrapper GetOrCreateWrapper(ServerCallContext context)
        {
            string sid = null;
            var key = context.Peer;
            if (_enable_session == true)
            {
                key = sid = context.RequestHeaders.FirstOrDefault(x => x.Key == "x-session-id")?.Value ?? Guid.NewGuid().ToString();
            }
            _logger.LogDebug("Request received. Peer:{0} SessionId:{1} ConnectionId:{2}", context.Peer, sid, context.GetHttpContext().Connection.Id);
            //foreach (var kv in context.RequestHeaders)
            //{
            //    _logger.LogDebug("header {0}: {1}", kv.Key, kv.Value);
            //}
            //foreach (var kv in context.GetHttpContext().Request.Cookies)
            //{
            //    _logger.LogDebug("cookie {0}: {1}", kv.Key, kv.Value);
            //}
            return _cache.GetOrCreate(key, entry =>
            {
                _logger.LogInformation("New connection comming! {0}", key);
                entry.AbsoluteExpirationRelativeToNow = TimeSpan.FromSeconds(_session_timeout_sec ?? 60);
                entry.PostEvictionCallbacks.Add(new PostEvictionCallbackRegistration { EvictionCallback = (k,v,r,s) => _DisposeCache(v as NPoiWrapper) });
                context.GetHttpContext().Response.Headers.Add("x-session-id", sid);
                //context.GetHttpContext().Response.Cookies.Append("SESSION_ID", sid);
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
                var ret = new BytesValue { Value = ByteString.CopyFrom(GetOrCreateWrapper(context).Download()) };
                _logger.LogTrace("  --------------------- Current memory usage: {0} bytes ---------------------", Environment.WorkingSet);
                return ret;
            });

        public override Task<RecalculationPolicy> GetRecalculationPolicy(Empty _, ServerCallContext context)
                  => Task.Factory.StartNew(() =>
                  {
                      return new RecalculationPolicy { Value = GetOrCreateWrapper(context).recalculationPolicy };
                  });

        public override Task<Empty> SetRecalculationPolicy(RecalculationPolicy p, ServerCallContext context)
                  => Task.Factory.StartNew(() =>
                  {
                      GetOrCreateWrapper(context).recalculationPolicy = p.Value;
                      return new Empty();
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

        public override Task<Empty> RemoveSheetAt(Int32Value arg, ServerCallContext context)
            => Task.Factory.StartNew(() =>
            {
                _logger.LogDebug("RemoveSheetAt({0})", arg);
                GetOrCreateWrapper(context).RemoveSheetAt(arg.Value);
                return new Empty();
            });

        public override Task<Empty> SetSheetHidden(IndexAndState arg, ServerCallContext context)
            => Task.Factory.StartNew(() =>
            {
                _logger.LogDebug("SetSheetHidden({0}, {1})", arg.Index, arg.State);
                GetOrCreateWrapper(context).SetSheetHidden(arg.Index, (NPOI.SS.UserModel.SheetState)arg.State);
                return new Empty();
            });
    
        public override Task<Empty> SetSheetName(IndexAndName arg, ServerCallContext context)
            => Task.Factory.StartNew(() =>
            {
                _logger.LogDebug("SetSheetName({0}, '{1}')", arg.Index, arg.Name);
                GetOrCreateWrapper(context).SetSheetName(arg.Index, arg.Name);
                return new Empty();
            });

        public override Task<Int32Value> CloneSheet(IndexAndName arg, ServerCallContext context)
            => Task.Factory.StartNew(() =>
            {
                _logger.LogDebug("CloneSheet({0}, '{1}')", arg.Index, arg.Name);
                var ret = GetOrCreateWrapper(context).CloneSheet(arg.Index, arg.Name);
                return new Int32Value { Value = ret };
            });
        public override Task<Empty> ClearRowAt(Int32Value arg, ServerCallContext context)
            => Task.Factory.StartNew(() =>
            {
                _logger.LogDebug("ClearRowAt({0})", arg);
                GetOrCreateWrapper(context).ClearRowAt(arg.Value);
                return new Empty();
            });

        public override Task<Empty> SetRowZeroHeighAt(IndexAndFlag arg, ServerCallContext context)
            => Task.Factory.StartNew(() =>
            {
                _logger.LogDebug("SetRowZeroHeight({0}, {1})", arg.Index, arg.Flag);
                GetOrCreateWrapper(context).SetRowZeroHeightAt(arg.Index, arg.Flag);
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
        public override Task<Int32Value> GetMaxRowNum(Empty _, ServerCallContext context)
            => Task.Factory.StartNew(() =>
            {
                _logger.LogDebug("GetMaxRowNum()");
                return new Int32Value { Value = GetOrCreateWrapper(context).GetMaxRowNum() };
            });
    }
}
