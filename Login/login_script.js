// login_script.js

let loginUsers = []; // Variabel global untuk menyimpan data pengguna yang diambil

document.addEventListener('DOMContentLoaded', async () => {
    const loginForm = document.getElementById('login-form');
    const loginMessage = document.getElementById('login-message');

    // *** PENTING: GANTI URL INI DENGAN URL WEB APP GOOGLE APPS SCRIPT ANDA ***
    const APPS_SCRIPT_LOGIN_DATA_URL = "https://script.google.com/macros/s/AKfycbxd3IuTELFcUm4H_XqBbpVyJrH5zQHzvpuk0rNKUslI1Jxj9MPfPKOiBNRQuu8qqeqG/exec"; // Contoh: "https://script.google.com/macros/s/AKfyc..."

    // Fungsi untuk mengambil data login dari Google Apps Script
    async function fetchLoginData() {
        try {
            loginMessage.textContent = 'Memuat data login...';
            loginMessage.className = 'login-message'; // Hapus styling error/success
            loginMessage.style.display = 'block';

            const response = await fetch(APPS_SCRIPT_LOGIN_DATA_URL);
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            loginUsers = await response.json();
            console.log("Data login berhasil dimuat:", loginUsers);
            loginMessage.style.display = 'none'; // Sembunyikan pesan jika berhasil
            return true;
        } catch (error) {
            console.error('Gagal memuat data login:', error);
            loginMessage.textContent = 'Kesalahan memuat data login. Silakan coba lagi nanti.';
            loginMessage.className = 'login-message error';
            loginMessage.style.display = 'block';
            return false;
        }
    }

    // Ambil data saat halaman dimuat
    const dataLoaded = await fetchLoginData();
    if (!dataLoaded) {
        // Nonaktifkan pengiriman formulir jika data gagal dimuat
        loginForm.querySelector('button[type="submit"]').disabled = true;
    }

    loginForm.addEventListener('submit', (e) => {
        e.preventDefault(); // Mencegah pengiriman formulir default

        const username = loginForm.username.value;
        const password = loginForm.password.value; // Ini sesuai dengan kolom 'Cabang' di sheet Anda

        // Cari pengguna di data yang diambil
        const foundUser = loginUsers.find(user =>
            user.Email === username && user.Cabang === password
        );

        if (foundUser) {
            loginMessage.textContent = 'Login berhasil! Mengarahkan...';
            loginMessage.className = 'login-message success';
            loginMessage.style.display = 'block';

            // Simpan status autentikasi di sessionStorage
            sessionStorage.setItem('authenticated', 'true');
            sessionStorage.setItem('loggedInUserEmail', username); // Opsional: simpan username
            sessionStorage.setItem('loggedInUserCabang', password); // Opsional: simpan cabang/password

            setTimeout(() => {
                window.location.href = 'index.html'; // Arahkan ke halaman landing utama Anda
            }, 1500);
        } else {
            loginMessage.textContent = 'Username atau password salah.';
            loginMessage.className = 'login-message error';
            loginMessage.style.display = 'block';
        }
    });
});