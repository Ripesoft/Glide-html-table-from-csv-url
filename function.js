window.function = async function(isDownload, hasHeader, fileUrl) {
  const extension = fileUrl.split('.').pop().toLowerCase();
  if (extension === 'csv' || extension === 'tsv') {
    return new Promise((resolve, reject) => {
      Papa.parse(fileUrl, {
        download: isDownload,
        header: hasHeader,
        complete: function(results) {
          resolve(results.data);
        },
        error: function(err) {
          reject(err);
        }
      });
    });
  } else if (extension === 'xls' || extension === 'xlsx') {
    try {
      const response = await fetch(fileUrl);
      if (!response.ok) {
        throw new Error('Failed to fetch file');
      }
      const arrayBuffer = await response.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer, { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const data = XLSX.utils.sheet_to_json(sheet, { header: hasHeader ? 1 : undefined });
      return data;
    } catch (err) {
      throw err;
    }
  } else {
    throw new Error('Unsupported file type. Supported: csv, tsv, xls, xlsx');
  }
};