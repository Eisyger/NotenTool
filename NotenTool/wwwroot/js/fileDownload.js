window.downloadFile = function(dataArray, fileName) {
    // dataArray kommt als Uint8Array von .NET
    var blob = new Blob([dataArray], { type: 'application/octet-stream' });
    var link = document.createElement('a');
    link.href = window.URL.createObjectURL(blob);
    link.download = fileName;
    link.click();

    // Cleanup - wichtig für Speicher
    window.URL.revokeObjectURL(link.href);
};
window.pickFile = async (accept) => {
    return new Promise((resolve, reject) => {
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = accept;

        input.onchange = async (e) => {
            const file = e.target.files[0];
            if (!file) {
                resolve(null);
                return;
            }

            try {
                const arrayBuffer = await file.arrayBuffer();
                const bytes = new Uint8Array(arrayBuffer);
                const base64 = btoa(String.fromCharCode(...bytes));

                // Objekt mit allen Infos zurückgeben
                resolve({
                    fileName: file.name,
                    lastModified: file.lastModified, // Timestamp in Millisekunden
                    size: file.size,
                    type: file.type,
                    content: base64
                });
            } catch (error) {
                console.error('Error reading file:', error);
                reject(error);
            }
        };

        input.oncancel = () => {
            resolve(null);
        };

        input.click();
    });
};