import { SmtpClient } from "https://deno.land/x/smtp@v0.7.0/mod.ts";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
};

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") {
    return new Response("ok", { headers: corsHeaders });
  }

  try {
    const { artista, obras } = await req.json();

    const listaArtigos = obras.map((o: { titulo: string; oferta: boolean; preco: string; quantidade: string }, i: number) => {
      const preco = o.oferta
        ? "Oferta"
        : `€${parseFloat(o.preco).toFixed(2)} × ${o.quantidade} ${parseInt(o.quantidade) === 1 ? "peça" : "peças"}`;
      return `  ${i + 1}. ${o.titulo} — ${preco}`;
    }).join("\n");

    const totalVenda = obras.filter((o: { oferta: boolean }) => !o.oferta).length;
    const totalOferta = obras.filter((o: { oferta: boolean }) => o.oferta).length;

    const corpo = [
      `Nova submissão de inventário recebida no portal Obscena.`,
      ``,
      `DADOS DO ARTISTA`,
      `Nome: ${artista.nome}`,
      `Nome artístico: ${artista.nomeArtistico || "—"}`,
      `Email: ${artista.email}`,
      `Contacto: ${artista.contacto}`,
      ``,
      `INVENTÁRIO — ${obras.length} ${obras.length === 1 ? "artigo" : "artigos"} (${totalVenda} para venda, ${totalOferta} para oferta)`,
      listaArtigos,
    ].join("\n");

    const client = new SmtpClient();
    await client.connectTLS({
      hostname: "smtp.gmail.com",
      port: 465,
      username: Deno.env.get("GMAIL_USER")!,
      password: Deno.env.get("GMAIL_APP_PASSWORD")!,
    });

    await client.send({
      from: Deno.env.get("GMAIL_USER")!,
      to: "a.malandra.producao@gmail.com",
      subject: "Nova submissão de inventário - Obscena",
      content: corpo,
    });

    await client.close();

    return new Response(JSON.stringify({ ok: true }), {
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  } catch (e) {
    console.error("Erro ao enviar email:", e);
    return new Response(JSON.stringify({ error: (e as Error).message }), {
      status: 500,
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }
});
