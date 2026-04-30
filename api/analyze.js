// api/analyze.js — Vercel Serverless Function
// Menggunakan Google Gemini API (gratis — gemini-2.0-flash)
// Env variable yang dibutuhkan: GEMINI_API_KEY

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST')   return res.status(405).json({ error: 'Method not allowed' });

  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    return res.status(500).json({
      error: 'GEMINI_API_KEY belum diatur. Buka Vercel → Settings → Environment Variables.'
    });
  }

  try {
    const { system, messages } = req.body;

    // Ambil konten dari messages (format Anthropic → konversi ke Gemini)
    const userMessage = messages?.[0]?.content || [];

    // Bangun parts untuk Gemini
    const parts = [];

    if (Array.isArray(userMessage)) {
      userMessage.forEach(item => {
        if (item.type === 'image') {
          // Gemini pakai inlineData untuk gambar base64
          parts.push({
            inlineData: {
              mimeType: item.source?.media_type || 'image/jpeg',
              data: item.source?.data || ''
            }
          });
        } else if (item.type === 'text') {
          parts.push({ text: item.text });
        }
      });
    } else if (typeof userMessage === 'string') {
      parts.push({ text: userMessage });
    }

    // Badan request Gemini
    const geminiBody = {
      systemInstruction: {
        parts: [{ text: system || '' }]
      },
      contents: [
        { role: 'user', parts }
      ],
      generationConfig: {
        maxOutputTokens: 1000,
        temperature: 0.2       // rendah agar output JSON konsisten
      }
    };

    // Panggil Gemini API — model gratis: gemini-2.0-flash
    const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${apiKey}`;

    const response = await fetch(endpoint, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(geminiBody)
    });

    const geminiData = await response.json();

    if (!response.ok) {
      const msg = geminiData?.error?.message || 'Gemini API error';
      return res.status(response.status).json({ error: msg });
    }

    // Ambil teks dari respons Gemini
    const text = geminiData?.candidates?.[0]?.content?.parts?.[0]?.text || '{}';

    // Konversi ke format respons yang sama dengan sebelumnya
    // agar frontend tidak perlu diubah
    return res.status(200).json({
      content: [{ type: 'text', text }]
    });

  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
}
