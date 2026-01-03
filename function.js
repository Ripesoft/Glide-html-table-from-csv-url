window.function = async function(isDownload, hasHeader, fileUrl) {
  // Determine filename or URL string safely (support File/Blob objects)
  let extension = '';
  if (typeof fileUrl === 'string') {
    extension = fileUrl.split('.').pop().toLowerCase();
  } else if (fileUrl && (fileUrl.name || fileUrl.fileName)) {
    const name = fileUrl.name || fileUrl.fileName;
    extension = name.split('.').pop().toLowerCase();
  }

  if (extension === 'csv' || extension === 'tsv') {
    return new Promise((resolve, reject) => {
      const config = {
        header: hasHeader,
        complete: function(results) {
          resolve(results.data);
        },
        error: function(err) {
          reject(err);
        }
      };

      // If a URL string was provided, allow downloading when requested
      if (typeof fileUrl === 'string') {
        config.download = isDownload;
        Papa.parse(fileUrl, config);
      } else {
        // For File/Blob objects, pass the object directly to Papa.parse
        Papa.parse(fileUrl, config);
      }
    });
  } else if (extension === 'xls' || extension === 'xlsx') {
    try {
      let arrayBuffer;

      if (typeof fileUrl === 'string') {
        const response = await fetch(fileUrl);
        if (!response.ok) {
          throw new Error('Failed to fetch file');
        }
        arrayBuffer = await response.arrayBuffer();
      } else {
        // For File/Blob objects, prefer the built-in arrayBuffer method
        if (typeof fileUrl.arrayBuffer === 'function') {
          arrayBuffer = await fileUrl.arrayBuffer();
        } else {
          // Fallback: use the Fetch API Response to obtain an ArrayBuffer
          arrayBuffer = await new Response(fileUrl).arrayBuffer();
        }
      }

      const workbook = XLSX.read(arrayBuffer, { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const data = XLSX.utils.sheet_to_json(sheet, { header: hasHeader ? 1 : undefined });
      return data;
    } catch (err) {
      throw err;
    }
  } else {
    return 'return: Unsupported file type. Supported: csv, tsv, xls, xlsx';
  }
};