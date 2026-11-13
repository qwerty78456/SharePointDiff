using System.IO.Compression;
using System.Net;
using System.Net.Sockets;
using System.Text;
using SharePointDiff.Core;
using Xunit;

namespace SharePointDiff.Tests;

public sealed class CompressionTests
{
    [Theory]
    [InlineData("gzip")]
    [InlineData("br")]
    [InlineData("deflate")]
    public async Task ProductionJsonHandlerNegotiatesAndDecodesCompressedBodies(string encoding)
    {
        string json = "{\"value\":[" + string.Join(',', Enumerable.Repeat("{\"name\":\"Tài liệu 中文\",\"Revision\":\"A\"}", 500)) + "]}";
        byte[] original = Encoding.UTF8.GetBytes(json);
        using var compressed = new MemoryStream();
        using (Stream compressor = encoding switch
        {
            "gzip" => new GZipStream(compressed, CompressionLevel.Fastest, true),
            "br" => new BrotliStream(compressed, CompressionLevel.Fastest, true),
            _ => new DeflateStream(compressed, CompressionLevel.Fastest, true)
        }) await compressor.WriteAsync(original);
        byte[] payload = compressed.ToArray();
        using var server = new TcpListener(IPAddress.Loopback, 0); server.Start();
        var address = (IPEndPoint)server.LocalEndpoint;
        string headers = "";
        var serving = Task.Run(async () =>
        {
            using var client = await server.AcceptTcpClientAsync();
            await using var connection = client.GetStream(); using var reader = new StreamReader(connection, Encoding.ASCII, false, 1024, true);
            string? line;
            while ((line = await reader.ReadLineAsync()) is { Length: > 0 }) headers += line + "\n";
            byte[] response = Encoding.ASCII.GetBytes($"HTTP/1.1 200 OK\r\nContent-Type: application/json\r\nContent-Encoding: {encoding}\r\nContent-Length: {payload.Length}\r\nConnection: close\r\n\r\n");
            await connection.WriteAsync(response); await connection.WriteAsync(payload); await connection.FlushAsync();
        });
        using var http = HttpTransport.Client(true);
        Assert.Equal(json, await http.GetStringAsync($"http://127.0.0.1:{address.Port}/metadata").WaitAsync(TimeSpan.FromSeconds(10)));
        await serving.WaitAsync(TimeSpan.FromSeconds(10));
        Assert.Contains("Accept-Encoding:", headers); Assert.Contains(encoding, headers);
        Assert.True(payload.Length < original.Length / 5);
    }
}
