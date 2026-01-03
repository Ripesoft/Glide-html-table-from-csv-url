window.function = async function(isDownload, hasHeader, fileUrl) {
  // Unwrap Glide parameter objects (they arrive as { value: ... })
  isDownload = isDownload?.value ?? isDownload;
  hasHeader = hasHeader?.value ?? hasHeader;
  const fileParam = fileUrl?.value ?? fileUrl;

  // Determine filename or URL string safely (support File/Blob objects)
  let extension = '';
  if (typeof fileParam === 'string') {
    extension = fileParam.split('.').pop().toLowerCase();
  } else if (fileParam && (fileParam.name || fileParam.fileName)) {
    const name = fileParam.name || fileParam.fileName;
    extension = name.split('.').pop().toLowerCase();
  }

  // Helper to return a JSON-stringified, consistent result envelope
  const okResult = (payload) => JSON.stringify({ ok: true, payload });
  const errResult = (message) => JSON.stringify({ ok: false, error: message });

  if (extension === 'csv' || extension === 'tsv') {
    return await new Promise((resolve) => {
      const config = {
        header: hasHeader,
        complete: function(results) {
          resolve(okResult({ format: extension, data: results.data }));
        },
        error: function(err) {
          resolve(errResult('CSV parse error: ' + (err?.message || err)));
        }
      };

      // If a URL string was provided, allow downloading when requested
      if (typeof fileParam === 'string') {
        config.download = isDownload;
        Papa.parse(fileParam, config);
      } else {
        // For File/Blob objects, pass the object directly to Papa.parse
        Papa.parse(fileParam, config);
      }
    });
  } else if (extension === 'xls' || extension === 'xlsx') {
    try {
      let arrayBuffer;

      if (typeof fileParam === 'string') {
        const response = await fetch(fileParam);
        if (!response.ok) {
          return errResult('Failed to fetch file: ' + response.status);
        }
        arrayBuffer = await response.arrayBuffer();
      } else {
        // For File/Blob objects, prefer the built-in arrayBuffer method
        if (typeof fileParam.arrayBuffer === 'function') {
          arrayBuffer = await fileParam.arrayBuffer();
        } else {
          // Fallback: use the Fetch API Response to obtain an ArrayBuffer
          arrayBuffer = await new Response(fileParam).arrayBuffer();
        }
      }

      const workbook = XLSX.read(arrayBuffer, { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const data = XLSX.utils.sheet_to_json(sheet, { header: hasHeader ? 1 : undefined });
      return okResult({ format: 'xlsx', sheetName, data });
    } catch (err) {
      return errResult('XLSX parse error: ' + (err?.message || err));
    }
  } else {
    return errResult('Unsupported file type. Supported: csv, tsv, xls, xlsx');
  }
};