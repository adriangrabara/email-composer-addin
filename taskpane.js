/* global Office */

// State
let selectedTone = 'formal';
let generatedEmail = '';
let isOfficeInitialized = false;

// Initialize Office.js
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    isOfficeInitialized = true;
    console.log('Office.js initialized for Outlook');
    
    // Try to get the original email if we're in a reply context
    tryGetOriginalEmail();
  }
  
  // Initialize the UI regardless
  initializeUI();
});

// Initialize UI elements and event listeners
function initializeUI() {
  // Initialize Lucide icons
  if (typeof lucide !== 'undefined') {
    lucide.createIcons();
  }
  
  // Collapsible section toggle
  const toggle = document.getElementById('originalEmailToggle');
  const content = document.getElementById('originalEmailContent');
  
  toggle.addEventListener('click', () => {
    toggle.classList.toggle('open');
    content.classList.toggle('open');
  });
  
  // Tone selector
  const toneButtons = document.querySelectorAll('.tone-btn');
  toneButtons.forEach(btn => {
    btn.addEventListener('click', () => {
      toneButtons.forEach(b => b.classList.remove('selected'));
      btn.classList.add('selected');
      selectedTone = btn.dataset.tone;
    });
  });
  
  // Generate button
  document.getElementById('generateBtn').addEventListener('click', generateEmail);
  
  // Copy button
  document.getElementById('copyBtn').addEventListener('click', copyToClipboard);
  
  // Insert button
  document.getElementById('insertBtn').addEventListener('click', insertIntoEmail);
}

// Try to get the original email content if replying
function tryGetOriginalEmail() {
  if (!isOfficeInitialized) return;
  
  try {
    const item = Office.context.mailbox.item;
    
    // Check if we're reading an email (could be replying to it)
    if (item && item.body) {
      item.body.getAsync(Office.CoercionType.Text, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded && result.value) {
          // Only populate if there's meaningful content
          const content = result.value.trim();
          if (content.length > 10) {
            document.getElementById('originalEmail').value = content;
            // Open the collapsible section
            document.getElementById('originalEmailToggle').classList.add('open');
            document.getElementById('originalEmailContent').classList.add('open');
          }
        }
      });
    }
  } catch (error) {
    console.log('Could not get original email:', error);
  }
}

// Generate email using Claude API
async function generateEmail() {
  const rawThoughts = document.getElementById('rawThoughts').value.trim();
  const originalEmail = document.getElementById('originalEmail').value.trim();
  const generateBtn = document.getElementById('generateBtn');
  const errorMsg = document.getElementById('errorMsg');
  const outputSection = document.getElementById('outputSection');
  
  // Validation
  if (!rawThoughts) {
    showError('Please enter your thoughts first');
    return;
  }
  
  // Hide any previous error
  hideError();
  
  // Show loading state
  setLoadingState(true);
  
  // Build the prompt
  const toneDescriptions = {
    formal: 'Formal - Professional and structured',
    warm: 'Warm - Friendly and approachable',
    concise: 'Concise - Brief and to the point',
    persuasive: 'Persuasive - Compelling and convincing',
    apologetic: 'Apologetic - Sincere and understanding',
    enthusiastic: 'Enthusiastic - Energetic and positive'
  };
  
  const prompt = `You are an expert email writer. Transform the following raw thoughts into a polished, professional email.

Tone: ${toneDescriptions[selectedTone]}

${originalEmail ? `This is a reply to the following email:
---
${originalEmail}
---

` : ''}Raw thoughts to transform:
${rawThoughts}

Write ONLY the email content. Do not include any explanations, meta-commentary, or notes. Start directly with the greeting (like "Hi [Name]," or "Dear [Name],") and end with a sign-off. Make sure the email:
- Matches the requested tone perfectly
- Is well-structured with clear paragraphs
- ${originalEmail ? 'Appropriately references and responds to the original email' : 'Has an appropriate subject line suggestion at the very start in the format "Subject: [subject]"'}
- Sounds natural and human
- Uses [Name] or [Your Name] as placeholders where appropriate`;

  try {
    const response = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        model: "claude-sonnet-4-20250514",
        max_tokens: 1000,
        messages: [
          { role: "user", content: prompt }
        ],
      })
    });

    const data = await response.json();
    
    if (data.content && data.content[0] && data.content[0].text) {
      generatedEmail = data.content[0].text;
      displayOutput(generatedEmail);
    } else if (data.error) {
      showError(data.error.message || 'Failed to generate email');
    } else {
      showError('Unexpected response format');
    }
  } catch (error) {
    console.error('API Error:', error);
    showError('Failed to connect to the API. Please check your connection and try again.');
  } finally {
    setLoadingState(false);
  }
}

// Display the generated email
function displayOutput(content) {
  const outputSection = document.getElementById('outputSection');
  const outputContent = document.getElementById('outputContent');
  
  outputContent.textContent = content;
  outputSection.classList.remove('hidden');
  
  // Re-initialize icons for the output section
  if (typeof lucide !== 'undefined') {
    lucide.createIcons();
  }
  
  // Scroll to output
  outputSection.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

// Copy to clipboard
async function copyToClipboard() {
  const copyBtn = document.getElementById('copyBtn');
  
  try {
    await navigator.clipboard.writeText(generatedEmail);
    showSuccess('Copied to clipboard!');
    
    // Temporarily update button
    copyBtn.innerHTML = '<i data-lucide="check"></i><span>Copied!</span>';
    copyBtn.classList.add('btn-success');
    
    if (typeof lucide !== 'undefined') {
      lucide.createIcons();
    }
    
    setTimeout(() => {
      copyBtn.innerHTML = '<i data-lucide="copy"></i><span>Copy</span>';
      copyBtn.classList.remove('btn-success');
      if (typeof lucide !== 'undefined') {
        lucide.createIcons();
      }
    }, 2000);
  } catch (error) {
    showError('Failed to copy to clipboard');
  }
}

// Insert into Outlook email
function insertIntoEmail() {
  if (!isOfficeInitialized) {
    // Fallback: just copy to clipboard
    copyToClipboard();
    showSuccess('Office not available - copied to clipboard instead');
    return;
  }
  
  try {
    const item = Office.context.mailbox.item;
    
    if (item && item.body) {
      // Check if we're in compose mode
      if (item.body.setAsync) {
        // Insert the email content
        item.body.setAsync(
          generatedEmail,
          { coercionType: Office.CoercionType.Text },
          (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              showSuccess('Email inserted successfully!');
            } else {
              // Try to prepend instead
              item.body.prependAsync(
                generatedEmail + '\n\n',
                { coercionType: Office.CoercionType.Text },
                (prependResult) => {
                  if (prependResult.status === Office.AsyncResultStatus.Succeeded) {
                    showSuccess('Email inserted successfully!');
                  } else {
                    showError('Could not insert email. Please copy and paste manually.');
                    copyToClipboard();
                  }
                }
              );
            }
          }
        );
      } else {
        // Not in compose mode - copy to clipboard
        copyToClipboard();
        showSuccess('Copied! Open a new email to paste.');
      }
    } else {
      copyToClipboard();
      showSuccess('Copied! Open a new email to paste.');
    }
  } catch (error) {
    console.error('Insert error:', error);
    copyToClipboard();
    showSuccess('Copied to clipboard - please paste into your email.');
  }
}

// UI Helper functions
function setLoadingState(isLoading) {
  const generateBtn = document.getElementById('generateBtn');
  
  if (isLoading) {
    generateBtn.disabled = true;
    generateBtn.innerHTML = '<i data-lucide="loader-2" class="spinner"></i><span>Crafting your email...</span>';
  } else {
    generateBtn.disabled = false;
    generateBtn.innerHTML = '<i data-lucide="send"></i><span>Generate Email</span>';
  }
  
  if (typeof lucide !== 'undefined') {
    lucide.createIcons();
  }
}

function showError(message) {
  const errorMsg = document.getElementById('errorMsg');
  errorMsg.textContent = message;
  errorMsg.classList.remove('hidden');
}

function hideError() {
  const errorMsg = document.getElementById('errorMsg');
  errorMsg.classList.add('hidden');
}

function showSuccess(message) {
  const status = document.getElementById('successStatus');
  status.querySelector('span').textContent = message;
  status.classList.remove('hidden');
  
  if (typeof lucide !== 'undefined') {
    lucide.createIcons();
  }
  
  setTimeout(() => {
    status.classList.add('hidden');
  }, 3000);
}
