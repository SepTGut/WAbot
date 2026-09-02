module.exports = {
  apps: [
    {
      name: "wa-bot",
      script: "index.js",
      instances: 1,
      exec_mode: "fork",
      autorestart: true,
      watch: false, // JANGAN di-true karena bot menulis ke file excel & json secara berkala
      max_memory_restart: "500M",
      restart_delay: 5000, // Jeda 5 detik sebelum restart jika crash
      env: {
        NODE_ENV: "production",
      },
    },
  ],
};
