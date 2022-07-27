let i;

WebViewer({
    path: '/assets/js/lib',
}, document.getElementById('viewer'))
    .then(async (instance) => {
        const filepath = window.location.search.replace('?filepath=', '');
        i = instance;
        const doc = await instance.Core.createDocument(filepath, { extension: 'pdf' });
        instance.loadDocument(doc);
    })

const saveButton = document.getElementById('save-button');
saveButton.addEventListener('click', async () => {
    let annotationManager = i.Core.annotationManager;
    const xfdfString = await annotationManager.exportAnnotations();
    const content = await i.Core.documentViewer.getDocument().getFileData({ xfdfString });
    const contentSize = await i.Core.documentViewer.getDocument().getFileSize();
    const isLargeFile = contentSize > 4 * 1024 * 1024;
    const tokenResp = await fetch('/token', {
        method: 'GET',
        headers: {
            'Content-Type': 'application/json'
        }
    });
    const tokenResult = await tokenResp.json();
    if (tokenResult.message !== 'success') {
        return alert("Couldn't retreive token.");
    }
    const { activationItems, token } = tokenResult.data;
    const itemResp = await fetch(activationItems, {
        headers: {
            'Authorization': `Bearer ${token}`,
        }
    });
    if (itemResp.status !== 200) {
        return alert("Failed to retreive the file information.")
    }
    const itemInfo = await itemResp.clone().json();
    const itemUrl = `https://graph.microsoft.com/v1.0/drives/${itemInfo.parentReference.driveId}/items/${itemInfo.id}`;

    if (isLargeFile) {
        uploadLargeFile(itemUrl, content, token);
    } else {
        uploadSmallFile(itemUrl, content, token);
    }
});

async function uploadLargeFile(itemUrl, content, token) {
    const uploadSessionUrl = `${itemUrl}/createUploadSession`;
    const uploadRequestResp = await fetch(uploadSessionUrl, {
        method: 'POST',
        headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({
            item: {
              "@microsoft.graph.conflictBehavior": "replace",
            },
        }),
    });
    const uploadRequestRespJson = await uploadRequestResp.json();
    if (uploadRequestRespJson.error) {
        return alert('upload request fails');
    }
    const { "uploadUrl": uploadUrl } = uploadRequestRespJson;
    const chunkLimit = 1024 * 1024 * 2;
    const totalSize = content.byteLength;
    let currentSizeOffset = 0;
    while(currentSizeOffset < totalSize) {
        const chunkSize = Math.min(chunkLimit, totalSize - currentSizeOffset);
        const chunk = content.slice(currentSizeOffset, currentSizeOffset + chunkSize);

        const uploadResp = await fetch(uploadUrl, {
            method: 'PUT',
            headers: {
                'Authorization': `Bearer ${token}`,
                'Content-Length': chunk.byteLength.toString(),
                'Content-Range': `bytes ${currentSizeOffset}-${
                    currentSizeOffset + chunkSize - 1
                }/${totalSize}`,
            },
            body: chunk
        });
        if (!uploadResp.ok) {
            return alert('upload failed')
        }
        currentSizeOffset += chunkSize;
    }
    return alert('Saved the file successfully!')
};

async function uploadSmallFile(itemUrl, content, token) {
    const contentUrl = `${itemUrl}/content`;
    const uploadResult = await fetch(contentUrl, {
        body: content,
        headers: {
        authorization: `Bearer ${token}`,
        },
        method: "PUT",
    });

    if (!uploadResult.ok) {
        return alert('Upload failed.')
    }
}