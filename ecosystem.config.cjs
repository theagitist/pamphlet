module.exports = {
  apps: [
    {
      name: 'pamphlet',
      script: './app.js',
      instances: 1,
      autorestart: true,
      watch: false,
      max_memory_restart: '500M',
      kill_timeout: 5000,
      env: {
        NODE_ENV: 'production',
        PORT: 3000,
        PAMPHLET_STORAGE_ROOT: '/mnt/polivoxiadata/pamphlet.polivoxia.ca',
        MAX_FILE_SIZE_MB: 50,
        CONCURRENCY_LIMIT: 3,
        MAX_QUEUE_DEPTH: 10,
        CLEANUP_MINUTES: 10,
      },
    },
  ],
};
