window.downloadFile = function (dataArray, fileName, mimeType) {

    if (!mimeType) {
        mimeType = "application/octet-stream";
    }

    const blob = new Blob([dataArray], { type: mimeType });
    const url = window.URL.createObjectURL(blob);

    const isIOS =
        /iPad|iPhone|iPod/.test(navigator.userAgent) ||
        (navigator.platform === "MacIntel" && navigator.maxTouchPoints > 1);

    if (isIOS) {
        // iOS: Blob direkt öffnen
        window.location.href = url;

        // Mehr Delay, Safari braucht Zeit
        setTimeout(() => {
            window.URL.revokeObjectURL(url);
        }, 8000);

        return;
    }

    // Desktop + Android
    const link = document.createElement("a");
    link.href = url;
    link.download = fileName;
    link.rel = "noopener";
    link.style.display = "none";

    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);

    setTimeout(() => {
        window.URL.revokeObjectURL(url);
    }, 1500);
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