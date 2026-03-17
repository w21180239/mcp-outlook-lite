function buildPage({ title, heading, headingColor, iconColor, svgPath, bodyText, instructionText, extraScript = '' }) {
  return `
    <html>
      <head>
        <title>${title}</title>
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
            color: ${headingColor};
            margin-bottom: 20px;
          }
          .icon {
            width: 80px;
            height: 80px;
            margin: 0 auto 20px;
            background-color: ${iconColor};
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
          }
          .icon svg {
            width: 50px;
            height: 50px;
            fill: white;
          }
          .info {
            color: #605e5c;
            font-size: 14px;
            margin-top: 20px;
          }
          .minor {
            font-size: 12px;
            color: #a19f9d;
            margin-top: 10px;
          }
        </style>
      </head>
      <body>
        <div class="container">
          <div class="icon">
            <svg viewBox="0 0 24 24">
              <path d="${svgPath}"/>
            </svg>
          </div>
          <h1>${heading}</h1>
          <p>${bodyText}</p>
          ${instructionText ? `<p class="info">${instructionText}</p>` : ''}
        </div>
        ${extraScript}
      </body>
    </html>
  `;
}

export function getSuccessPage() {
  return buildPage({
    title: 'Authentication Successful',
    heading: 'Authentication Successful!',
    headingColor: '#0078d4',
    iconColor: '#107c10',
    svgPath: 'M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41L9 16.17z',
    bodyText: 'The Outlook MCP server has been configured with your selected account.',
    instructionText: 'This window will close in <span id="timer">5</span> seconds...',
    extraScript: `
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
                document.querySelector('.info').textContent = 'You can now close this window.';
              }, 500);
            }
          }, 1000);
        </script>
    `,
  });
}

export function getErrorPage() {
  return buildPage({
    title: 'Authentication Error',
    heading: 'Security Error',
    headingColor: '#d83b01',
    iconColor: '#d83b01',
    svgPath: 'M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 15h-2v-2h2v2zm0-4h-2V7h2v6z',
    bodyText: 'The authentication request could not be verified.',
    instructionText: 'Please disconnect and reconnect the MCP server to try again.',
  });
}

export function getFailurePage() {
  return buildPage({
    title: 'Authentication Failed',
    heading: 'Authentication Failed',
    headingColor: '#d83b01',
    iconColor: '#d83b01',
    svgPath: 'M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12 19 6.41z',
    bodyText: 'The authentication process was cancelled or failed.',
    instructionText: 'Please disconnect and reconnect the MCP server to try again.',
  });
}
