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

  // Helper to convert data array to HTML table
  const dataToHtmlTable = (data, sheetName = null) => {
    if (!data || !Array.isArray(data) || data.length === 0) {
      return '<table style="border-collapse: collapse; width: 100%;"><tr><td>No data</td></tr></table>';
    }

    let html = '<table style="border-collapse: collapse; width: 100%; border: 1px solid #ddd;">';
    let keys;
    let startIdx = 0;

    // If data is array of objects, use object keys
    if (typeof data[0] === 'object' && !Array.isArray(data[0])) {
      keys = Object.keys(data[0]);
      if (hasHeader) {
        html += '<thead><tr style="background-color: #f2f2f2;">';
        keys.forEach(key => {
          html += `<th style="border: 1px solid #ddd; padding: 8px; text-align: left;">${key}</th>`;
        });
        html += '</tr></thead>';
      }
    } else if (Array.isArray(data[0])) {
      // If data is array of arrays, use first row as header if hasHeader
      keys = data[0].map((_, idx) => idx);
      if (hasHeader) {
        html += '<thead><tr style="background-color: #f2f2f2;">';
        data[0].forEach(cell => {
          html += `<th style="border: 1px solid #ddd; padding: 8px; text-align: left;">${cell}</th>`;
        });
        html += '</tr></thead>';
        startIdx = 1;
      }
    }

    // Add data rows
    html += '<tbody>';
    data.forEach((row, rowIdx) => {
      if (Array.isArray(row) && hasHeader && rowIdx === 0) return; // skip header row for array-of-arrays
      html += `<tr style="${rowIdx % 2 === 0 ? 'background-color: #ffffff;' : 'background-color: #f9f9f9;'}">`;
      if (typeof row === 'object' && !Array.isArray(row)) {
        keys.forEach(key => {
          html += `<td style="border: 1px solid #ddd; padding: 8px;">${row[key] ?? ''}</td>`;
        });
      } else if (Array.isArray(row)) {
        row.forEach(cell => {
          html += `<td style="border: 1px solid #ddd; padding: 8px;">${cell ?? ''}</td>`;
        });
      }
      html += '</tr>';
    });
    html += '</tbody></table>';

    return html;
  };

  // Helper to generate tabbed HTML for multiple sheets
  const generateTabbedHtml = (sheetData) => {
    // sheetData is an array of { sheetName, data }
    if (!Array.isArray(sheetData) || sheetData.length === 0) {
      return '<div>No data available</div>';
    }

    let html = `
      <div style="font-family: Arial, sans-serif;">
        <div style="display: flex; border-bottom: 2px solid #ddd; margin-bottom: 10px; align-items: flex-end;">
    `;

    // Create minimalistic tabs with black and gray only
    sheetData.forEach((sheet, idx) => {
      const isActive = idx === 0 ? 'true' : 'false';
      html += `
        <button onclick="document.querySelectorAll('[data-sheet-content]').forEach(el => el.style.display='none'); document.querySelector('[data-sheet-content='${sheet.sheetName}']').style.display='block'; document.querySelectorAll('[data-sheet-tab]').forEach(el => {el.style.borderBottom='2px solid #eee'; el.style.color='#666';}); this.style.borderBottom='2px solid #111'; this.style.color='#111';"
                data-sheet-tab="${sheet.sheetName}"
                style="padding: 6px 0 0 0; min-width: 0; flex: 1 1 0; text-align: left; cursor: pointer; background: none; border: none; outline: none; font-size: 15px; transition: color 0.2s, border-bottom 0.2s; color: ${isActive === 'true' ? '#111' : '#666'}; border-bottom: 2px solid ${isActive === 'true' ? '#111' : '#eee'}; margin-right: 8px;">
          ${sheet.sheetName}
        </button>
      `;
    });

    html += '</div>';

    // Create content divs for each sheet
    sheetData.forEach((sheet, idx) => {
      html += `
        <div data-sheet-content="${sheet.sheetName}" style="display: ${idx === 0 ? 'block' : 'none'};">
          ${dataToHtmlTable(sheet.data, sheet.sheetName)}
        </div>
      `;
    });

    html += '</div>';
    return html;
  };

  if (extension === 'csv' || extension === 'tsv') {
    return await new Promise((resolve) => {
      const config = {
        header: hasHeader,
        complete: function(results) {
          const htmlTable = dataToHtmlTable(results.data);
          resolve(htmlTable);
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
      
      // Process all sheets and create tabbed HTML
      const sheetData = workbook.SheetNames.map(sheetName => {
        const sheet = workbook.Sheets[sheetName];
        // If hasHeader, let XLSX infer headers (objects); else, get array of arrays
        let data;
        if (hasHeader) {
          data = XLSX.utils.sheet_to_json(sheet, { header: undefined });
        } else {
          data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        }
        return { sheetName, data };
      });

      const html = generateTabbedHtml(sheetData);
      
      return html;
    } catch (err) {
      return errResult('XLSX parse error: ' + (err?.message || err));
    }
  } else {
    return errResult('Unsupported file type. Supported: csv, tsv, xls, xlsx');
  }
};