module.exports = {
    apps: [
      {
        name: "ocr_cv_invoice",
        script: "./venv/bin/streamlit", // Path to the Node.js wrapper script
        args: ['run', 'Home.py'],
        instances: 1,
        autorestart: true,
        watch: false,
        interpreter: './venv/bin/python',
        max_memory_restart: "1G",
      }
    ]
  };
