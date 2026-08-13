// scanner.js — mendukung scanner fisik (keyboard/HID) dan kamera (html5-qrcode)

// ---------- SCANNER FISIK ----------
// Scanner fisik "mengetik" karakter lalu Enter secara otomatis.
// Kita cukup dengar event Enter pada input yang sedang fokus.
function initPhysicalScanner(inputElement, onScanCallback) {
    inputElement.addEventListener('keydown', function (e) {
        if (e.key === 'Enter') {
            e.preventDefault();
            const kode = inputElement.value.trim();
            if (kode.length > 0) {
                onScanCallback(kode);
                inputElement.value = '';
            }
        }
    });
    // auto-focus biar siap nangkep scan tanpa perlu klik dulu
    inputElement.focus();
}

// ---------- SCANNER KAMERA ----------
// Dipakai html5-qrcode. Pastikan modal HTML dengan id yang sesuai sudah ada di halaman.
let html5QrcodeScanner = null;

function startCameraScanner(modalId, readerId, onScanCallback) {
    const modalEl = document.getElementById(modalId);
    const modal = new bootstrap.Modal(modalEl);
    modal.show();

    html5QrcodeScanner = new Html5Qrcode(readerId);

    Html5Qrcode.getCameras().then(cameras => {
        if (cameras && cameras.length) {
            const cameraId = cameras[cameras.length - 1].id;
            html5QrcodeScanner.start(
                cameraId,
                {
                    fps: 10,
                    qrbox: { width: 300, height: 150 },
                    experimentalFeatures: {
                        useBarCodeDetectorIfSupported: true
                    },  // dilebarkan, barcode 1D butuh area lebih lebar dari tinggi
                    aspectRatio: 1.7777778,
                    formatsToSupport: [
                        Html5QrcodeSupportedFormats.EAN_13,
                        Html5QrcodeSupportedFormats.EAN_8,
                        Html5QrcodeSupportedFormats.CODE_128,
                        Html5QrcodeSupportedFormats.QR_CODE
                    ],
                    videoConstraints: {
                        facingMode: "environment",
                        focusMode: "continuous",
                        width: { ideal: 1920 },
                        height: { ideal: 1080 },
                        advanced: [{ zoom: 2 }]
                    }
                },
                (decodedText) => {
                    stopCameraScanner(modal);
                    onScanCallback(decodedText);
                },
                (errorMessage) => { /* diabaikan */ }
            );
        } else {
            alert('Kamera tidak ditemukan di perangkat ini.');
            modal.hide();
        }
    }).catch(err => {
        alert('Gagal mengakses kamera: ' + err);
        modal.hide();
    });

    modalEl.addEventListener('hidden.bs.modal', () => stopCameraScanner(modal), { once: true });
}

function stopCameraScanner(modal) {
    if (html5QrcodeScanner) {
        html5QrcodeScanner.stop().then(() => {
            html5QrcodeScanner.clear();
            html5QrcodeScanner = null;
        }).catch(() => {});
    }
    if (modal) modal.hide();
}