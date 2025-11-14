// Utility function to generate Recommendation Letter Word document template
// Creates Word-compatible HTML document that can be opened in Microsoft Word

export interface RecommendationLetterData {
  projectName: string;
  client: string;
  location: string;
  completionDate: string;
  poNumber: string;
  manager: string;
  clientContact?: string;
}

// Generate Word document for recommendation letter
export const generateRecommendationLetterWord = async (data: RecommendationLetterData): Promise<Blob> => {
  try {
    // Create Word-compatible HTML document
    // Microsoft Word can open HTML files saved with .doc extension
    const htmlContent = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>Recommendation Letter Template</title>
  <style>
    body {
      font-family: 'Times New Roman', serif;
      font-size: 12pt;
      line-height: 1.6;
      margin: 1in;
      color: #000;
    }
    .header {
      text-align: center;
      margin-bottom: 30px;
    }
    .header h1 {
      font-size: 18pt;
      font-weight: bold;
      margin-bottom: 10px;
    }
    .section {
      margin-bottom: 20px;
    }
    .section-title {
      font-size: 14pt;
      font-weight: bold;
      margin-bottom: 10px;
      border-bottom: 1px solid #000;
      padding-bottom: 5px;
    }
    .field {
      margin-bottom: 15px;
    }
    .field-label {
      font-weight: bold;
      margin-bottom: 5px;
    }
    .field-input {
      border-bottom: 1px solid #000;
      min-height: 20px;
      padding: 5px;
      margin-top: 5px;
    }
    .signature-section {
      margin-top: 50px;
      margin-bottom: 30px;
    }
    .signature-line {
      border-top: 1px solid #000;
      width: 300px;
      margin-top: 50px;
    }
    .project-details {
      background-color: #f5f5f5;
      padding: 15px;
      border: 1px solid #ddd;
      margin-bottom: 20px;
    }
    .project-details p {
      margin: 5px 0;
    }
  </style>
</head>
<body>
  <div class="header">
    <h1>RECOMMENDATION LETTER</h1>
  </div>

  <div class="project-details">
    <p><strong>Project Name:</strong> ${data.projectName}</p>
    <p><strong>Client:</strong> ${data.client}</p>
    <p><strong>Location:</strong> ${data.location}</p>
    <p><strong>Completion Date:</strong> ${data.completionDate}</p>
    <p><strong>PO Number:</strong> ${data.poNumber}</p>
    <p><strong>Project Manager:</strong> ${data.manager}</p>
  </div>

  <div class="section">
    <div class="section-title">Company Details</div>
    <div class="field">
      <div class="field-label">Company Name:</div>
      <div class="field-input">&nbsp;</div>
    </div>
    <div class="field">
      <div class="field-label">Company Address:</div>
      <div class="field-input">&nbsp;</div>
    </div>
    <div class="field">
      <div class="field-label">Contact Person:</div>
      <div class="field-input">${data.clientContact || '________________'}</div>
    </div>
    <div class="field">
      <div class="field-label">Email:</div>
      <div class="field-input">&nbsp;</div>
    </div>
    <div class="field">
      <div class="field-label">Phone:</div>
      <div class="field-input">&nbsp;</div>
    </div>
  </div>

  <div class="section">
    <div class="section-title">Recommendation & Feedback</div>
    <div class="field">
      <div class="field-label">Please provide your feedback on the quality of work delivered:</div>
      <div class="field-input" style="min-height: 60px;">&nbsp;</div>
    </div>
    <div class="field">
      <div class="field-label">Please comment on adherence to timelines and specifications:</div>
      <div class="field-input" style="min-height: 60px;">&nbsp;</div>
    </div>
    <div class="field">
      <div class="field-label">Please provide feedback on professional conduct and communication:</div>
      <div class="field-input" style="min-height: 60px;">&nbsp;</div>
    </div>
    <div class="field">
      <div class="field-label">Overall satisfaction and recommendation:</div>
      <div class="field-input" style="min-height: 80px;">&nbsp;</div>
    </div>
    <div class="field">
      <div class="field-label">Additional Comments (if any):</div>
      <div class="field-input" style="min-height: 60px;">&nbsp;</div>
    </div>
  </div>

  <div class="signature-section">
    <div class="field">
      <div class="field-label">Authorized Signatory Name:</div>
      <div class="field-input">&nbsp;</div>
    </div>
    <div class="field">
      <div class="field-label">Designation:</div>
      <div class="field-input">&nbsp;</div>
    </div>
    <div class="field">
      <div class="field-label">Date:</div>
      <div class="field-input">&nbsp;</div>
    </div>
    <div class="signature-line"></div>
    <div style="margin-top: 10px;">Signature & Company Seal</div>
  </div>

  <div style="margin-top: 30px; font-size: 10pt; color: #666;">
    <p><em>Please fill in all the details, sign this document, and send it back to us. Thank you for your time and feedback.</em></p>
  </div>
</body>
</html>
    `;

    // Convert HTML to Word document format
    // Add UTF-8 BOM for proper character encoding in Word
    // Use application/msword MIME type for .doc files
    const blob = new Blob(['\ufeff', htmlContent], {
      type: 'application/msword;charset=utf-8'
    });

    return blob;
  } catch (error) {
    console.error('Error generating Word document:', error);
    throw error;
  }
};

// Download the generated Word file (kept for backward compatibility, but not used in new flow)
export const downloadWordFile = (blob: Blob, filename: string) => {
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
};
