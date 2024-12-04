let jsonData = []; // Global variable to hold the imported data

    document.getElementById('fileInput').addEventListener('change', function () {
      const fileInput = document.getElementById('fileInput');
      const importedDataDiv = document.getElementById('importedData');

      if (!fileInput.files[0]) {
        alert('Bitte wählen Sie eine Excel-Datei aus!');
        return;
      }

      const reader = new FileReader();
      reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];

        // Parse sheet to JSON with headers
        jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        // Extract headers
        const headers = jsonData[0] || []; // First row as headers or empty array if none
        const requiredHeaders = ['name', 'email', 'pin'];
        const missingHeaders = requiredHeaders.filter(header => !headers.includes(header));

        // Check for missing headers
        if (missingHeaders.length > 0) {
          alert(`Die Excel-Datei fehlt folgende Spaltenüberschrift(en): ${missingHeaders.join(', ')}`);
          jsonData = []; // Clear jsonData on invalid input
          return;
        }

        // Remove header row from jsonData and map to objects
        jsonData = jsonData.slice(1).map(row => {
          return {
            name: row[headers.indexOf('name')],
            email: row[headers.indexOf('email')],
            pin: row[headers.indexOf('pin')],
          };
        });

        // Display imported data
        importedDataDiv.classList.remove('alert-info');
        importedDataDiv.classList.add('alert-success');
        importedDataDiv.innerHTML = '<ul class="list-group">';
        jsonData.forEach((row, index) => {
          importedDataDiv.innerHTML += `<li class="list-group-item">Name: ${row.name}, Email: ${row.email}, PIN: ${row.pin}</li>`;
        });
        importedDataDiv.innerHTML += '</ul>';
      };

      reader.readAsArrayBuffer(fileInput.files[0]);
    });

    // Export configuration to a file
    document.getElementById('exportConfig').addEventListener('click', function () {
      const config = {
        emailSubject: document.getElementById('emailSubject').value,
        emailBody: document.getElementById('emailBody').value,
        smtpServer: document.getElementById('smtpServer').value,
        smtpPort: document.getElementById('smtpPort').value,
        smtpUser: document.getElementById('smtpUser').value,
        smtpPass: document.getElementById('smtpPass').value
      };

      const blob = new Blob([JSON.stringify(config, null, 2)], { type: 'application/json' });
      const link = document.createElement('a');
      link.href = URL.createObjectURL(blob);
      link.download = 'config.json';
      link.click();
    });

    // Import configuration from a file
    document.getElementById('importConfig').addEventListener('click', function () {
      const input = document.createElement('input');
      input.type = 'file';
      input.accept = '.json';

      input.addEventListener('change', function (e) {
        const file = e.target.files[0];
        const reader = new FileReader();

        reader.onload = function (e) {
          try {
            const config = JSON.parse(e.target.result);

            // Fill the fields with the imported configuration
            document.getElementById('emailSubject').value = config.emailSubject || '';
            document.getElementById('emailBody').value = config.emailBody || '';
            document.getElementById('smtpServer').value = config.smtpServer || '';
            document.getElementById('smtpPort').value = config.smtpPort || '465';
            document.getElementById('smtpUser').value = config.smtpUser || '';
            document.getElementById('smtpPass').value = config.smtpPass || '';
          } catch (err) {
            alert('Fehler beim Importieren der Konfiguration.');
          }
        };

        reader.readAsText(file);
      });

      input.click();
    });

    document.getElementById('generateEmails').addEventListener('click', function () {
      const emailSubject = document.getElementById('emailSubject').value.trim();
      const emailBodyTemplate = document.getElementById('emailBody').value;
      const emailListDiv = document.getElementById('emailList');

      // Warnung anzeigen, wenn das Betreff-Feld leer ist
      if (!emailSubject) {
        alert('Bitte geben Sie einen Betreff für die Emails ein!');
        return;
      }

      emailListDiv.innerHTML = ''; // Clear previous results

      // Validate if data is imported
      if (jsonData.length === 0) {
        alert('Bitte laden Sie zuerst eine gültige Excel-Datei hoch!');
        return;
      }

      let delay = 0; // Start delay at 0ms
      jsonData.forEach((row) => {
        const email = row.email;
        const name = row.name;
        const pin = row.pin;

        if (!email || !name || !pin) return; // Skip invalid rows

        setTimeout(() => {
          const personalizedBody = emailBodyTemplate
            .replace('{{name}}', name)
            .replace('{{pin}}', pin);

          const mailtoLink = `mailto:${email}?subject=${encodeURIComponent(emailSubject)}&body=${encodeURIComponent(personalizedBody)}`;
          const emailElement = document.createElement('a');
          emailElement.href = mailtoLink;
          emailElement.textContent = `Email an ${email}`;
          emailElement.target = '_blank';
          emailListDiv.appendChild(emailElement);
          emailListDiv.appendChild(document.createElement('br'));

          // Automatically open the email in Outlook
          const hiddenLink = document.createElement('a');
          hiddenLink.href = mailtoLink;
          hiddenLink.style.display = 'none';
          document.body.appendChild(hiddenLink);
          hiddenLink.click();
          document.body.removeChild(hiddenLink);
        }, delay);

        delay += 500; // Increment delay for the next email
      });
    });