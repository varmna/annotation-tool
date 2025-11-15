// Initialize main variables
const tool = {
    currentIndex: 0,
    conversations: [],
    annotations: {},
    buckets: [
        "Bot Response",
        "HVA",
        "AB feature/HVA Related Query",
        "Personalized/Account-Specific Queries",
        "Promo & Freebie Related Queries",
        "Help-page/Direct Customer Service",
        "BP for Non-Profit Organisation Related Query",
        "Personal Prime Related Query",
        "Customer Behavior",
        "Other Queries",
        "Overall Observations"
    ]
};

// Initialize UI elements
const elements = {
    uploadScreen: document.getElementById('upload-screen'),
    toolInterface: document.getElementById('to'tool-interface'),
    fileInput: document.getElementById('excel-upload'),
    conversationDisplay: document.getElementById('conversation-display'),
    conversationList: document.getElementById('conversation-list'),
    bucketArea: document.getElementById('bucket-area'),
    prevBtn: document.getElementById('prev-btn'),
    nextBtn: document.getElementById('next-btn'),
    saveBtn: document.getElementById('save-btn'),
    downloadBtn: document.getElementById('download-btn'),
    progress: document.getElementById('progress'),
    progressText: document.getElementById('progress-text'),
    statusMessage: document.getElementById('status-message'),
    loadingSpinner: document.getElementById('loading-spinner'),
    currentTitle: document.getElementById('current-conversation-title')
};

// Create bucket UI
function createBucketUI() {
    tool.buckets.forEach(bucket => {
        const bucketHTML = `
            <div class="bucket">
                <label class="bucket-label">
                    <input type="checkbox" name="${bucket}">
                    <span>${bucket}</span>
                </label>
                <textarea 
                    placeholder="Add comments for ${bucket}" 
                    name="${bucket}"
                    rows="3"
                ></textarea>
            </div>
        `;
        elements.bucketArea.insertAdjacentHTML('beforeend', bucketHTML);
    });
}

// File upload handler
elements.fileInput.addEventListener('change', async (event) => {
    const file = event.target.files[0];
    if (!file) return;

    showLoading(true);
    showStatus('📂 Loading file...', 'info');

    try {
        const data = await readExcelFile(file);
        processExcelData(data);
        elements.uploadScreen.style.display = 'none';
        elements.toolInterface.style.display = 'flex';
        showStatus('✅ File loaded successfully!', 'success');
    } catch (error) {
        console.error('Error:', error);
        showStatus('❌ Error loading file', 'error');
    } finally {
        showLoading(false);
    }
});

// Read Excel file
async function readExcelFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, {type: 'array'});
                const sheetName = workbook.SheetNames[0];
                const sheet = workbook.Sheets[sheetName];
                resolve(XLSX.utils.sheet_to_json(sheet));
            } catch (error) {
                reject(error);
            }
        };
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

// Process Excel data
function processExcelData(rawData) {
    const groupedData = {};
    rawData.forEach(row => {
        if (!groupedData[row.Id]) {
            groupedData[row.Id] = [];
        }
        groupedData[row.Id].push(row);
    });
    
    tool.conversations = Object.values(groupedData);
    tool.currentIndex = 0;
    tool.annotations = {};
    
    updateConversationList();
    updateProgressBar();
    displayConversation();
}

// Create conversation list items
function updateConversationList() {
    elements.conversationList.innerHTML = '';
    tool.conversations.forEach((conv, index) => {
        const item = document.createElement('div');
        item.className = `conversation-item ${index === tool.currentIndex ? 'active' : ''}`;
        item.innerHTML = `
            <div>Conversation ${index + 1}</div>
            <small>ID: ${conv[0].Id}</small>
        `;
        item.onclick = () => {
            tool.currentIndex = index;
            updateConversationList();
            displayConversation();
        };
        elements.conversationList.appendChild(item);
    });
}

// Display conversation
function displayConversation() {
    const conv = tool.conversations[tool.currentIndex];
    const lastMessage = conv[conv.length - 1];

    elements.currentTitle.textContent = `Conversation ${tool.currentIndex + 1} of ${tool.conversations.length}`;

    let html = `
        <div class="conversation-info">
            <strong>ID:</strong> ${conv[0].Id}<br>
            <strong>Customer Feedback:</strong> 
            <span class="badge ${lastMessage['Customer Feedback']?.toLowerCase() === 'negative' ? 'bg-danger' : 'bg-success'}">
                ${lastMessage['Customer Feedback'] || 'N/A'}
            </span>
        </div>
        <div class="messages">
    `;

    conv.forEach(message => {
        if (message.llmGeneratedUserMessage) {
            html += `
                <div class="message customer">
                    <div class="message-header">👤 Customer</div>
                    ${message.llmGeneratedUserMessage}
                </div>
            `;
        }
        if (message.botMessage) {
            html += `
                <div class="message bot">
                    <div class="message-header">🤖 Bot</div>
                    ${message.botMessage}
                </div>
            `;
        }
    });

    html += '</div>';
    elements.conversationDisplay.innerHTML = html;
    updateProgressBar();
    loadAnnotations();
}

// Update progress bar
function updateProgressBar() {
    const progress = ((tool.currentIndex + 1) / tool.conversations.length) * 100;
    elements.progress.style.width = `${progress}%`;
    elements.progressText.textContent = 
        `${tool.currentIndex + 1}/${tool.conversations.length} Conversations`;
}

// Save annotations
function saveCurrentAnnotations() {
    const convId = tool.conversations[tool.currentIndex][0].Id;
    const hasAnnotations = tool.buckets.some(bucket => 
        document.querySelector(`input[name="${bucket}"]`).checked
    );

    if (!hasAnnotations) {
        showStatus('⚠️ Please select at least one bucket', 'warning');
        return;
    }

    tool.annotations[convId] = {};
    
    tool.buckets.forEach(bucket => {
        const checkbox = document.querySelector(`input[name="${bucket}"]`);
        const textarea = document.querySelector(`textarea[name="${bucket}"]`);
        if (checkbox.checked) {
            tool.annotations[convId][bucket] = textarea.value.trim();
        }
    });

    showStatus('✅ Annotations saved!', 'success');
    updateConversationList();
}

// Load annotations
function loadAnnotations() {
    const convId = tool.conversations[tool.currentIndex][0].Id;
    const savedAnnotations = tool.annotations[convId] || {};
    
    tool.buckets.forEach(bucket => {
        const checkbox = document.querySelector(`input[name="${bucket}"]`);
        const textarea = document.querySelector(`textarea[name="${bucket}"]`);
        checkbox.checked = false;
        textarea.value = '';
    });

    Object.entries(savedAnnotations).forEach(([bucket, comment]) => {
        const checkbox = document.querySelector(`input[name="${bucket}"]`);
        const textarea = document.querySelector(`textarea[name="${bucket}"]`);
        if (checkbox && textarea) {
            checkbox.checked = true;
            textarea.value = comment;
        }
    });
}

// Show status message
function showStatus(message, type) {
    elements.statusMessage.textContent = message;
    elements.statusMessage.className = `status-message alert alert-${type}`;
    elements.statusMessage.style.display = 'block';
    
    setTimeout(() => {
        elements.statusMessage.style.display = 'none';
    }, 3000);
}

// Show/hide loading spinner
function showLoading(show) {
    elements.loadingSpinner.style.display = show ? 'flex' : 'none';
}

// Helper function for Excel binary conversion
function s2ab(s) {
    const buf = new ArrayBuffer(s.length);
    const view = new Uint8Array(buf);
    for (let i = 0; i < s.length; i++) {
        view[i] = s.charCodeAt(i) & 0xFF;
    }
    return buf;
}

// Download annotations
elements.downloadBtn.addEventListener('click', () => {
    try {
        if (Object.keys(tool.annotations).length === 0) {
            showStatus('⚠️ No annotations to download', 'warning');
            return;
        }

        showLoading(true);
        showStatus('💾 Preparing download...', 'info');
        
        const annotatedData = [];
        
        tool.conversations.forEach(conv => {
            const convId = conv[0].Id;
            const savedAnnotations = tool.annotations[convId];
            
            if (savedAnnotations && Object.keys(savedAnnotations).length > 0) {
                conv.forEach((message, index) => {
                    const isFirstMessage = index === 0;
                    const isLastMessage = index === conv.length - 1;
                    
                    const row = {
                        'Id': message.Id,
                        'llmGeneratedUserMessage': message.llmGeneratedUserMessage || '',
                        'botMessage': message.botMessage || '',
                        'Customer Feedback': isLastMessage ? message['Customer Feedback'] || '' : ''
                    };

                    if (isFirstMessage) {
                        tool.buckets.forEach(bucket => {
                            row[bucket] = savedAnnotations[bucket] || '';
                        });
                    } else {
                        tool.buckets.forEach(bucket => {
                            row[bucket] = '';
                        });
                    }
                    
                    annotatedData.push(row);
                });
            }
        });

        const ws = XLSX.utils.json_to_sheet(annotatedData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Annotations");

        const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'binary' });
        const blob = new Blob([s2ab(wbout)], { type: 'application/octet-stream' });
        const url = window.URL.createObjectURL(blob);

        const a = document.createElement('a');
        a.href = url;
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
        a.download = `annotated_conversations_${timestamp}.xlsx`;
        document.body.appendChild(a);
        a.click();
        
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);

        const annotatedCount = new Set(annotatedData.map(row => row.Id)).size;
        showStatus(`✅ Downloaded ${annotatedCount} conversation(s)!`, 'success');
    } catch (error) {
        console.error('Download error:', error);
        showStatus('❌ Error downloading file', 'error');
    } finally {
        showLoading(false);
    }
});

// Navigation handlers
elements.prevBtn.addEventListener('click', () => {
    if (tool.currentIndex > 0) {
        tool.currentIndex--;
        updateConversationList();
        displayConversation();
    } else {
        showStatus('⚠️ This is the first conversation', 'warning');
    }
});

elements.nextBtn.addEventListener('click', () => {
    if (tool.currentIndex < tool.conversations.length - 1) {
        tool.currentIndex++;
        updateConversationList();
        displayConversation();
    } else {
        showStatus('⚠️ This is the last conversation', 'warning');
    }
});

elements.saveBtn.addEventListener('click', saveCurrentAnnotations);

// Initialize buckets
createBucketUI();
