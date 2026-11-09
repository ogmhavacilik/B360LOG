import express from 'express';
import { createServer as createViteServer } from 'vite';

async function startServer() {
  const app = express();

  // CORS middleware for all routes
  app.use((req, res, next) => {
    res.header('Access-Control-Allow-Origin', '*');
    res.header('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
    res.header('Access-Control-Allow-Headers', 'Content-Type, Authorization');
    if (req.method === 'OPTIONS') {
      return res.sendStatus(200);
    }
    next();
  });

  app.use(express.json({ limit: '10mb' }));
  app.use(express.urlencoded({ extended: true, limit: '10mb' }));

  // Google Apps Script URL
  const GAS_URL = 'https://script.google.com/macros/s/AKfycbyEeDha6Fw4pviggDXAsUmQ1pEWm_TwTYDrirGZHKJxagOI8ZXLPEkLGgqyb9eNfhj0/exec';

  // Health check endpoint
  app.get('/api/health', (req, res) => {
    res.json({ status: 'ok', time: new Date().toISOString() });
  });

  // General proxy function for Google Apps Script interactions
  const proxyToGAS = async (req: express.Request, res: express.Response) => {
    const action = req.body?.action || (req.path.includes('delete') ? 'delete' : req.path.includes('update') ? 'update' : 'add');
    const enrichedBody = { ...req.body, action };

    console.log(`[GAS Proxy] ${req.method} ${req.path} -> action: ${action}, body:`, JSON.stringify(enrichedBody));

    try {
      const response = await fetch(GAS_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(enrichedBody),
        redirect: 'follow'
      });

      const responseText = await response.text();
      console.log(`[GAS Proxy] Response status: ${response.status}, body length: ${responseText.length}`);

      const isHtml = responseText.trim().startsWith('<') || responseText.includes('<!DOCTYPE') || responseText.includes('<html');

      if (isHtml) {
        // Extract plain text message from HTML error if returned by Google Sheets / Apps Script
        let extractedError = 'Google Apps Script hata döndürdü.';
        const divMatch = responseText.match(/<div[^>]*>([^<]{5,250})<\/div>/i);
        if (divMatch && divMatch[1]) {
          extractedError = divMatch[1].trim();
        } else {
          const titleMatch = responseText.match(/<title>([^<]+)<\/title>/i);
          if (titleMatch && titleMatch[1]) {
            extractedError = `GAS: ${titleMatch[1].trim()}`;
          }
        }
        console.error('[GAS Proxy] GAS returned HTML error:', extractedError);
        return res.status(400).json({ 
          success: false, 
          error: extractedError, 
          isHtml: true,
          details: responseText.slice(0, 500) 
        });
      }

      let parsedData: any = null;
      try {
        parsedData = JSON.parse(responseText);
      } catch (parseErr) {
        parsedData = { raw: responseText };
      }

      if (response.ok) {
        if (parsedData && (parsedData.status === 'error' || parsedData.success === false)) {
          console.warn('[GAS Proxy] GAS returned error status:', parsedData);
          return res.status(400).json({ success: false, error: parsedData.message || parsedData.error || 'GAS error', details: parsedData });
        }
        res.json({ success: true, action, data: parsedData });
      } else {
        res.status(response.status).json({ 
          success: false,
          error: parsedData?.message || 'GAS error', 
          status: response.status, 
          details: parsedData 
        });
      }
    } catch (error: any) {
      console.error('[GAS Proxy] Error forwarding to GAS:', error);
      res.status(500).json({ error: 'Failed to proxy request to Google Apps Script', message: error?.message });
    }
  };

  // Dedicated API Routes for Add, Update, Delete
  app.post('/api/add', proxyToGAS);
  app.post('/api/update', proxyToGAS);
  app.post('/api/delete', proxyToGAS);

  // Vite integration: Mount Vite middlewares AFTER API routes so API routes are not intercepted
  const vite = await createViteServer({
    server: { middlewareMode: true },
    appType: 'spa'
  });

  app.use(vite.middlewares);

  const PORT = 3000;
  app.listen(PORT, () => {
    console.log(`Server running at http://localhost:${PORT}`);
  });
}

startServer();
