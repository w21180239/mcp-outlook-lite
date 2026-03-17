export function getSuccessPage() {
  return `
    <html>
      <head>
        <title>Authentication Successful</title>
        <style>
          body {
            font-family: 'Segoe UI', Arial, sans-serif;
            text-align: center;
            padding: 50px;
            background-color: #f3f2f1;
            margin: 0;
          }
          .container {
            background: white;
            border-radius: 8px;
            padding: 40px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            max-width: 400px;
            margin: 0 auto;
          }
          h1 {
            color: #0078d4;
            margin-bottom: 20px;
          }
          .checkmark {
            width: 80px;
            height: 80px;
            margin: 0 auto 20px;
            background-color: #107c10;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            animation: scaleIn 0.3s ease-in-out;
          }
          .checkmark svg {
            width: 50px;
            height: 50px;
            fill: white;
          }
          @keyframes scaleIn {
            from { transform: scale(0); opacity: 0; }
            to { transform: scale(1); opacity: 1; }
          }
          .countdown {
            color: #605e5c;
            font-size: 14px;
            margin-top: 20px;
          }
          #timer {
            font-weight: bold;
            color: #0078d4;
          }
          .manual-close {
            font-size: 12px;
            color: #a19f9d;
            margin-top: 10px;
          }
        </style>
      </head>
      <body>
        <div class="container">
          <div class="checkmark">
            <svg viewBox="0 0 24 24">
              <path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41L9 16.17z"/>
            </svg>
          </div>
          <h1>Authentication Successful!</h1>
          <p>The Outlook MCP server has been configured with your selected account.</p>
          <p class="countdown">This window will close in <span id="timer">5</span> seconds...</p>
          <p class="manual-close">If the window doesn't close automatically, you can close it manually.</p>
        </div>
        <script>
          let countdown = 5;
          const timerElement = document.getElementById('timer');

          const interval = setInterval(() => {
            countdown--;
            timerElement.textContent = countdown;

            if (countdown <= 0) {
              clearInterval(interval);
              window.close();
              setTimeout(() => {
                document.querySelector('.countdown').textContent = 'You can now close this window.';
              }, 500);
            }
          }, 1000);
        </script>
      </body>
    </html>
  `;
}

export function getErrorPage() {
  return `
    <html>
      <head>
        <title>Authentication Error</title>
        <style>
          body {
            font-family: 'Segoe UI', Arial, sans-serif;
            text-align: center;
            padding: 50px;
            background-color: #f3f2f1;
            margin: 0;
          }
          .container {
            background: white;
            border-radius: 8px;
            padding: 40px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            max-width: 400px;
            margin: 0 auto;
          }
          h1 {
            color: #d83b01;
            margin-bottom: 20px;
          }
          .error-icon {
            width: 80px;
            height: 80px;
            margin: 0 auto 20px;
            background-color: #d83b01;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
          }
          .error-icon svg {
            width: 50px;
            height: 50px;
            fill: white;
          }
          .instructions {
            color: #605e5c;
            font-size: 14px;
            margin-top: 20px;
          }
        </style>
      </head>
      <body>
        <div class="container">
          <div class="error-icon">
            <svg viewBox="0 0 24 24">
              <path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 15h-2v-2h2v2zm0-4h-2V7h2v6z"/>
            </svg>
          </div>
          <h1>Security Error</h1>
          <p>The authentication request could not be verified.</p>
          <p class="instructions">Please disconnect and reconnect the MCP server to try again.</p>
        </div>
      </body>
    </html>
  `;
}

export function getFailurePage() {
  return `
    <html>
      <head>
        <title>Authentication Failed</title>
        <style>
          body {
            font-family: 'Segoe UI', Arial, sans-serif;
            text-align: center;
            padding: 50px;
            background-color: #f3f2f1;
            margin: 0;
          }
          .container {
            background: white;
            border-radius: 8px;
            padding: 40px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            max-width: 400px;
            margin: 0 auto;
          }
          h1 {
            color: #d83b01;
            margin-bottom: 20px;
          }
          .error-icon {
            width: 80px;
            height: 80px;
            margin: 0 auto 20px;
            background-color: #d83b01;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
          }
          .error-icon svg {
            width: 50px;
            height: 50px;
            fill: white;
          }
          .instructions {
            color: #605e5c;
            font-size: 14px;
            margin-top: 20px;
          }
        </style>
      </head>
      <body>
        <div class="container">
          <div class="error-icon">
            <svg viewBox="0 0 24 24">
              <path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12 19 6.41z"/>
            </svg>
          </div>
          <h1>Authentication Failed</h1>
          <p>The authentication process was cancelled or failed.</p>
          <p class="instructions">Please disconnect and reconnect the MCP server to try again.</p>
        </div>
      </body>
    </html>
  `;
}
