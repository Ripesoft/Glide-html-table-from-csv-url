window.function = function(isDownload, hasHeader, fileUrl) {
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
};