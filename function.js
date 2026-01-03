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
    
    // Get all keys from the first row to determine columns
    const firstRow = data[0];
    const keys = Array.isArray(firstRow) ? 
      firstRow.map((_, idx) => idx) : 
      Object.keys(firstRow);

    // Add header row
    if (hasHeader) {
      html += '<thead><tr style="background-color: #f2f2f2;">';
      keys.forEach(key => {
        html += `<th style="border: 1px solid #ddd; padding: 8px; text-align: left;">${key}</th>`;
      });
      html += '</tr></thead>';
    }

    // Add data rows
    html += '<tbody>';
    data.forEach((row, rowIdx) => {
      html += `<tr style="${rowIdx % 2 === 0 ? 'background-color: #ffffff;' : 'background-color: #f9f9f9;'}">`;
      keys.forEach(key => {
        const value = Array.isArray(row) ? row[key] : row[key];
        html += `<td style="border: 1px solid #ddd; padding: 8px;">${value ?? ''}</td>`;
      });
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
        <div style="display: flex; border-bottom: 2px solid #ddd; margin-bottom: 10px;">
    `;

    // Create tabs
    sheetData.forEach((sheet, idx) => {
      const isActive = idx === 0 ? 'true' : 'false';
      html += `
        <button onclick="document.querySelectorAll('[data-sheet-content]').forEach(el => el.style.display='none'); document.querySelector('[data-sheet-content=\"${sheet.sheetName}\"]').style.display='block'; document.querySelectorAll('[data-sheet-tab]').forEach(el => el.style.borderBottom=''); this.style.borderBottom='3px solid #4CAF50';" 
                data-sheet-tab="${sheet.sheetName}"
                style="padding: 10px 20px; cursor: pointer; background: none; border: none; font-size: 14px; ${isActive === 'true' ? 'border-bottom: 3px solid #4CAF50; font-weight: bold;' : ''}">
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
        const data = XLSX.utils.sheet_to_json(sheet, { header: hasHeader ? 1 : undefined });
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