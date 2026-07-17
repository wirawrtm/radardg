export async function onRequest(context) {
  const { request } = context;
  const url = new URL(request.url);

  // Ganti target URL ke backend Cloud Run Anda
  const targetHost = "https://ais-pre-5wvrk62z7vgpuz2pnct5zf-961275344911.asia-southeast1.run.app";
  const targetUrl = new URL(url.pathname + url.search, targetHost);

  // Buat request baru dengan header asli dan target URL yang baru
  const newRequest = new Request(targetUrl.toString(), {
    method: request.method,
    headers: request.headers,
    body: request.body,
    redirect: request.redirect
  });

  try {
    const response = await fetch(newRequest);
    
    // Kembalikan response dari backend Cloud Run ke client
    return response;
  } catch (error) {
    return new Response(JSON.stringify({
      status: "error",
      message: "Gagal menghubungkan ke backend Cloud Run melalui Cloudflare proxy: " + error.message
    }), {
      status: 502,
      headers: {
        "Content-Type": "application/json",
        "Access-Control-Allow-Origin": "*"
      }
    });
  }
}
