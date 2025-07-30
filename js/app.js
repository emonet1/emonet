// 全局常量
const CURRENT_TIME = '2025-03-16 11:48:09';
const CURRENT_USER = 'jingtianwei2002';
const ADMIN_USERNAME = 'admin';
const ADMIN_PASSWORD = 'admin123';

// Firebase 配置
const firebaseConfig = {
    apiKey: "AIzaSyBCoiNOTvbD_MjgdFLVaieeN1rq4F1hJaM",
    authDomain: "emonet-registration-system.firebaseapp.com",
    projectId: "emonet-registration-system",
    storageBucket: "emonet-registration-system.firebasestorage.app",
    messagingSenderId: "78485270665",
    appId: "1:78485270665:web:3927049f025372a3dfd658",
    measurementId: "G-LNPHXZH1C4"
};

// 初始化Firebase
try {
    firebase.initializeApp(firebaseConfig);
    console.log('Firebase 初始化成功');
} catch (error) {
    console.error('Firebase 初始化失败:', error);
    alert('系统初始化失败，请刷新页面重试');
}

const db = firebase.firestore();
let currentUser = null;
let currentExcelData = null;

// 工具函数
function showLoading(show = true) {
    const loader = document.getElementById('loadingIndicator');
    if (loader) {
        loader.style.display = show ? 'block' : 'none';
    }
}

function showError(elementId, message) {
    const element = document.getElementById(elementId);
    if (element) {
        const errorDiv = element.nextElementSibling;
        if (errorDiv && errorDiv.classList.contains('error-message')) {
            errorDiv.textContent = message;
            errorDiv.style.display = 'block';
        }
    }
}

function clearError(elementId) {
    const element = document.getElementById(elementId);
    if (element) {
        const errorDiv = element.nextElementSibling;
        if (errorDiv && errorDiv.classList.contains('error-message')) {
            errorDiv.style.display = 'none';
        }
    }
}

// 表单验证
function validateUsername(username) {
    return username.length > 0;
}

function validatePhone(phone) {
    return /^1[3-9]\d{9}$/.test(phone);
}

function validatePassword(password) {
    return password.length >= 6;
}

// 页面切换
function hideAllForms() {
    const forms = ['loginForm', 'registerForm', 'statusQueryForm', 'userPanel', 'adminPanel'];
    forms.forEach(id => {
        const element = document.getElementById(id);
        if (element) {
            element.style.display = 'none';
        }
    });
}

function showLoginForm() {
    hideAllForms();
    const loginForm = document.getElementById('loginForm');
    if (loginForm) {
        loginForm.style.display = 'block';
    }
}

function showRegisterForm() {
    hideAllForms();
    const registerForm = document.getElementById('registerForm');
    if (registerForm) {
        registerForm.style.display = 'block';
    }
}

function showStatusQuery() {
    hideAllForms();
    const statusQueryForm = document.getElementById('statusQueryForm');
    if (statusQueryForm) {
        statusQueryForm.style.display = 'block';
    }
}

// 登录处理
async function handleLogin(event) {
    event.preventDefault();
    console.log('处理登录请求...');

    const username = document.getElementById('loginUsername').value.trim();
    const password = document.getElementById('loginPassword').value;

    // 验证输入
    let isValid = true;
    if (!username) {
        showError('loginUsername', '请输入用户名');
        isValid = false;
    }
    if (!password) {
        showError('loginPassword', '请输入密码');
        isValid = false;
    }
    if (!isValid) return;

    showLoading(true);
    try {
        // 管理员登录
        if (username === ADMIN_USERNAME && password === ADMIN_PASSWORD) {
            console.log('管理员登录成功');
            currentUser = {
                username: ADMIN_USERNAME,
                role: 'admin',
                loginTime: CURRENT_TIME
            };
            hideAllForms();
            const adminPanel = document.getElementById('adminPanel');
            if (adminPanel) {
                adminPanel.style.display = 'block';
                await loadPendingUsers();
            }
            return;
        }

        // 普通用户登录
        const querySnapshot = await db.collection('users')
            .where('username', '==', username)
            .get();

        if (querySnapshot.empty) {
            showError('loginUsername', '用户名不存在');
            return;
        }

        const userDoc = querySnapshot.docs[0];
        const userData = userDoc.data();

        if (userData.password !== password) {
            showError('loginPassword', '密码错误');
            return;
        }

        if (userData.status !== 'approved') {
            showError('loginUsername', '账号尚未通过审核');
            return;
        }

        currentUser = {
            ...userData,
            id: userDoc.id
        };

        hideAllForms();
        const userPanel = document.getElementById('userPanel');
        const userWelcome = document.getElementById('userWelcome');
        if (userPanel && userWelcome) {
            userPanel.style.display = 'block';
            userWelcome.textContent = username;
            await loadUserCodeList();
        }

    } catch (error) {
        console.error('登录失败:', error);
        alert('登录失败: ' + error.message);
    } finally {
        showLoading(false);
    }
}

// 注册处理
async function handleRegister(event) {
    event.preventDefault();

    const username = document.getElementById('registerUsername').value.trim();
    const name = document.getElementById('registerName').value.trim();
    const company = document.getElementById('registerCompany').value.trim();
    const phone = document.getElementById('registerPhone').value.trim();
    const password = document.getElementById('registerPassword').value;

    // 验证输入
    let isValid = true;
    if (!validateUsername(username)) {
        showError('registerUsername', '用户名不能为空');
        isValid = false;
    }
    if (!name) {
        showError('registerName', '请输入姓名');
        isValid = false;
    }
    if (!company) {
        showError('registerCompany', '请输入单位名称');
        isValid = false;
    }
    if (!validatePhone(phone)) {
        showError('registerPhone', '请输入有效的手机号码');
        isValid = false;
    }
    if (!validatePassword(password)) {
        showError('registerPassword', '密码长度至少为6个字符');
        isValid = false;
    }
    if (!isValid) return;

    showLoading(true);
    try {
        // 检查用户名是否已存在
        const existingUser = await db.collection('users')
            .where('username', '==', username)
            .get();

        if (!existingUser.empty) {
            showError('registerUsername', '用户名已被使用');
            return;
        }

        // 创建新用户
        await db.collection('users').add({
            username,
            name,
            company,
            phone,
            password,
            status: 'pending',
            createdAt: CURRENT_TIME,
            createdBy: CURRENT_USER
        });

        alert('注册申请已提交，请等待管理员审核');
        showLoginForm();

    } catch (error) {
        console.error('注册失败:', error);
        alert('注册失败: ' + error.message);
    } finally {
        showLoading(false);
    }
}

// 状态查询
async function handleStatusQuery(event) {
    event.preventDefault();

    const phone = document.getElementById('queryPhone').value.trim();
    if (!validatePhone(phone)) {
        showError('queryPhone', '请输入有效的手机号码');
        return;
    }

    showLoading(true);
    try {
        const querySnapshot = await db.collection('users')
            .where('phone', '==', phone)
            .get();

        const resultDiv = document.getElementById('queryResult');
        if (!resultDiv) return;

        if (querySnapshot.empty) {
            resultDiv.innerHTML = `
                <div class="empty-state">
                    <i class="ri-search-line" style="font-size: 48px; color: #666;"></i>
                    <p>未找到注册信息</p>
                </div>
            `;
            return;
        }

        const userData = querySnapshot.docs[0].data();
        const statusText = {
            pending: '待审核',
            approved: '已通过',
            rejected: '已拒绝'
        }[userData.status];

        const statusClass = {
            pending: 'status-pending',
            approved: 'status-approved',
            rejected: 'status-rejected'
        }[userData.status];

        resultDiv.innerHTML = `
            <div class="card result-card">
                <h3>查询结果</h3>
                <div class="result-content">
                    <p><strong>用户名:</strong> ${userData.username}</p>
                    <p><strong>姓名:</strong> ${userData.name}</p>
                    <p><strong>单位:</strong> ${userData.company}</p>
                    <p><strong>状态:</strong> <span class="${statusClass}">${statusText}</span></p>
                    <p><strong>申请时间:</strong> ${userData.createdAt}</p>
                </div>
            </div>
        `;

    } catch (error) {
        console.error('查询失败:', error);
        alert('查询失败: ' + error.message);
    } finally {
        showLoading(false);
    }
}

// 管理员功能
async function loadPendingUsers() {
    showLoading(true);
    try {
        const snapshot = await db.collection('users')
            .where('status', '==', 'pending')
            .get();

        const pendingList = document.getElementById('pendingList');
        if (!pendingList) return;

        if (snapshot.empty) {
            pendingList.innerHTML = `
                <div class="empty-state">
                    <i class="ri-user-follow-line" style="font-size: 48px; color: #666;"></i>
                    <p>暂无待审核用户</p>
                </div>
            `;
            return;
        }

        pendingList.innerHTML = snapshot.docs.map(doc => {
            const user = doc.data();
            return `
                <div class="card user-card">
                    <div class="user-card-header">
                        <i class="ri-user-line user-icon"></i>
                        <h3>${user.username}</h3>
                    </div>
                    <div class="user-card-content">
                        <p><i class="ri-user-line"></i> 姓名: ${user.name}</p>
                        <p><i class="ri-building-line"></i> 单位: ${user.company}</p>
                        <p><i class="ri-phone-line"></i> 电话: ${user.phone}</p>
                        <p><i class="ri-time-line"></i> 申请时间: ${user.createdAt}</p>
                    </div>
                    <div class="user-card-actions">
                        <button onclick="handleApprove('${doc.id}')" class="btn btn-primary">
                            <i class="ri-check-line"></i> 通过
                        </button>
                        <button onclick="handleReject('${doc.id}')" class="btn btn-danger">
                            <i class="ri-close-line"></i> 拒绝
                        </button>
                    </div>
                </div>
            `;
        }).join('');
        
    } catch (error) {
        console.error('加载待审核用户失败:', error);
        alert('加载失败: ' + error.message);
    } finally {
        showLoading(false);
    }
}

async function handleApprove(userId) {
    if (!confirm('确定通过该用户的申请吗？')) return;

    showLoading(true);
    try {
        await db.collection('users').doc(userId).update({
            status: 'approved',
            approvedAt: CURRENT_TIME,
            approvedBy: currentUser.username
        });

        alert('已通过用户申请');
        await loadPendingUsers();

    } catch (error) {
        console.error('审批失败:', error);
        alert('操作失败: ' + error.message);
    } finally {
        showLoading(false);
    }
}

async function handleReject(userId) {
    if (!confirm('确定拒绝该用户的申请吗？')) return;

    showLoading(true);
    try {
        await db.collection('users').doc(userId).update({
            status: 'rejected',
            rejectedAt: CURRENT_TIME,
            rejectedBy: currentUser.username
        });

        alert('已拒绝用户申请');
        await loadPendingUsers();

    } catch (error) {
        console.error('拒绝失败:', error);
        alert('操作失败: ' + error.message);
    } finally {
        showLoading(false);
    }
}

// 代码文件管理
function showUploadForm() {
    const modal = document.getElementById('codeModal');
    const modalTitle = document.getElementById('modalTitle');
    const codeId = document.getElementById('codeId');
    const fileName = document.getElementById('fileName');
    const codeContent = document.getElementById('codeContent');

    if (modal && modalTitle && codeId && fileName && codeContent) {
        modalTitle.textContent = '上传代码';
        codeId.value = '';
        fileName.value = '';
        fileName.readOnly = false;
        codeContent.value = '';
        codeContent.readOnly = false;
        modal.style.display = 'block';
    }
}

function hideCodeModal() {
    const modal = document.getElementById('codeModal');
    if (modal) {
        modal.style.display = 'none';
    }
}

function getLanguageIcon(fileName) {
    const extension = fileName.split('.').pop().toLowerCase();
    const icons = {
        'js': '<i class="ri-javascript-line"></i>',
        'py': '<i class="ri-python-line"></i>',
        'java': '<i class="ri-code-line"></i>',
        'html': '<i class="ri-html5-line"></i>',
        'css': '<i class="ri-css3-line"></i>',
        'default': '<i class="ri-file-code-line"></i>'
    };
    return icons[extension] || icons.default;
}

function getCodePreview(content) {
    const maxLength = 150;
    const preview = content.slice(0, maxLength);
    return `<pre class="code-preview-content">${preview}${content.length > maxLength ? '...' : ''}</pre>`;
}

function formatDate(dateString) {
    try {
        const date = new Date(dateString);
        return date.toLocaleString('zh-CN', {
            year: 'numeric',
            month: '2-digit',
            day: '2-digit',
            hour: '2-digit',
            minute: '2-digit',
            second: '2-digit'
        });
    } catch (error) {
        return dateString;
    }
}

async function handleCodeSubmit(event) {
    event.preventDefault();

    const fileName = document.getElementById('fileName').value.trim();
    const content = document.getElementById('codeContent').value.trim();
    const codeId = document.getElementById('codeId').value;

    if (!fileName || !content) {
        alert('请填写所有必填字段');
        return;
    }

    showLoading(true);
    try {
        if (codeId) {
            // 更新现有代码
            await db.collection('codes').doc(codeId).update({
                fileName,
                content,
                updatedAt: CURRENT_TIME,
                updatedBy: currentUser.username
            });
        } else {
            // 添加新代码
            await db.collection('codes').add({
                fileName,
                content,
                createdAt: CURRENT_TIME,
                createdBy: currentUser.username,
                updatedAt: CURRENT_TIME,
                updatedBy: currentUser.username
            });
        }

        alert(codeId ? '代码已更新' : '代码已上传');
        hideCodeModal();
        
        if (currentUser.role === 'admin') {
            await loadAdminCodeList();
        } else {
            await loadUserCodeList();
        }

    } catch (error) {
        console.error('保存代码失败:', error);
        alert('操作失败: ' + error.message);
    } finally {
        showLoading(false);
    }
}

async function loadUserCodeList() {
    showLoading(true);
    try {
        const snapshot = await db.collection('codes')
            .orderBy('createdAt', 'desc')
            .get();

        const codeList = document.getElementById('userCodeList');
        if (!codeList) return;

        if (snapshot.empty) {
            codeList.innerHTML = `
                <div class="empty-state">
                    <i class="ri-code-box-line" style="font-size: 48px; color: #666;"></i>
                    <p>暂无代码文件</p>
                </div>
            `;
            return;
        }

        codeList.innerHTML = `
            <div class="code-list-header">
                <h2>代码文件列表</h2>
                <div class="code-list-actions">
                    <input type="text" id="searchCode" class="search-input" placeholder="搜索代码文件...">
                    <select id="languageFilter" class="filter-select">
                        <option value="">所有语言</option>
                        <option value="javascript">JavaScript</option>
                        <option value="python">Python</option>
                        <option value="java">Java</option>
                        <option value="html">HTML</option>
                        <option value="css">CSS</option>
                    </select>
                </div>
            </div>
            <div class="code-grid">
                ${snapshot.docs.map(doc => {
                    const code = doc.data();
                    return `
                        <div class="code-card" data-language="${getFileLanguage(code.fileName)}">
                            <div class="code-card-header">
                                <div class="code-icon">
                                    ${getLanguageIcon(code.fileName)}
                                </div>
                                <div class="code-info">
                                    <h3 class="code-title">${code.fileName}</h3>
                                    <span class="code-meta">创建者: ${code.createdBy}</span>
                                </div>
                            </div>
                            <div class="code-card-content">
                                <div class="code-preview">
                                    ${getCodePreview(code.content)}
                                </div>
                            </div>
                            <div class="code-card-footer">
                                <span class="code-date">创建时间: ${formatDate(code.createdAt)}</span>
                                <div class="code-actions">
                                    <button onclick="viewCode('${doc.id}')" class="btn btn-icon" title="查看">
                                        <i class="ri-eye-line"></i>
                                    </button>
                                </div>
                            </div>
                        </div>
                    `;
                }).join('')}
            </div>
        `;

        initializeCodeListFilters();

    } catch (error) {
        console.error('加载代码列表失败:', error);
        alert('加载失败: ' + error.message);
    } finally {
        showLoading(false);
    }
}

async function loadAdminCodeList() {
    showLoading(true);
    try {
        const snapshot = await db.collection('codes')
            .orderBy('createdAt', 'desc')
            .get();

        const codeList = document.getElementById('adminCodeList');
        if (!codeList) return;

        if (snapshot.empty) {
            codeList.innerHTML = `
                <div class="empty-state">
                    <i class="ri-code-box-line" style="font-size: 48px; color: #666;"></i>
                    <p>暂无代码文件</p>
                </div>
            `;
            return;
        }

        codeList.innerHTML = `
            <div class="code-list-header">
                <h2>代码文件管理</h2>
                <div class="code-list-actions">
                    <input type="text" id="adminSearchCode" class="search-input" placeholder="搜索代码文件...">
                    <select id="adminLanguageFilter" class="filter-select">
                        <option value="">所有语言</option>
                        <option value="javascript">JavaScript</option>
                        <option value="python">Python</option>
                        <option value="java">Java</option>
                        <option value="html">HTML</option>
                        <option value="css">CSS</option>
                    </select>
                    <button onclick="showUploadForm()" class="btn btn-primary">
                        <i class="ri-add-line"></i> 新建代码
                    </button>
                </div>
            </div>
            <div class="code-grid">
                ${snapshot.docs.map(doc => {
                    const code = doc.data();
                    return `
                        <div class="code-card" data-language="${getFileLanguage(code.fileName)}">
                            <div class="code-card-header">
                                <div class="code-icon">
                                    ${getLanguageIcon(code.fileName)}
                                </div>
                                <div class="code-info">
                                    <h3 class="code-title">${code.fileName}</h3>
                                    <span class="code-meta">创建者: ${code.createdBy}</span>
                                </div>
                            </div>
                            <div class="code-card-content">
                                <div class="code-preview">
                                    ${getCodePreview(code.content)}
                                </div>
                            </div>
                            <div class="code-card-footer">
                                <span class="code-date">更新时间: ${formatDate(code.updatedAt)}</span>
                                <div class="code-actions">
                                    <button onclick="viewCode('${doc.id}')" class="btn btn-icon" title="查看">
                                        <i class="ri-eye-line"></i>
                                    </button>
                                    <button onclick="editCode('${doc.id}')" class="btn btn-icon" title="编辑">
                                        <i class="ri-edit-line"></i>
                                    </button>
                                    <button onclick="deleteCode('${doc.id}')" class="btn btn-icon" title="删除">
                                        <i class="ri-delete-bin-line"></i>
                                    </button>
                                </div>
                            </div>
                        </div>
                    `;
                }).join('')}
            </div>
        `;

        initializeCodeListFilters('admin');

    } catch (error) {
        console.error('加载代码列表失败:', error);
        alert('加载失败: ' + error.message);
    } finally {
        showLoading(false);
    }
}

function initializeCodeListFilters(prefix = '') {
    const searchInput = document.getElementById(`${prefix}SearchCode`);
    const languageFilter = document.getElementById(`${prefix}LanguageFilter`);
    
    if (searchInput) {
        searchInput.addEventListener('input', (e) => filterCodeCards(prefix));
    }
    
    if (languageFilter) {
        languageFilter.addEventListener('change', (e) => filterCodeCards(prefix));
    }
}

function filterCodeCards(prefix = '') {
    const searchInput = document.getElementById(`${prefix}SearchCode`);
    const languageFilter = document.getElementById(`${prefix}LanguageFilter`);
    const codeCards = document.querySelectorAll('.code-card');

    const searchText = searchInput ? searchInput.value.toLowerCase() : '';
    const selectedLanguage = languageFilter ? languageFilter.value.toLowerCase() : '';

    codeCards.forEach(card => {
        const title = card.querySelector('.code-title').textContent.toLowerCase();
        const language = card.dataset.language.toLowerCase();
        
        const matchesSearch = title.includes(searchText);
        const matchesLanguage = !selectedLanguage || language === selectedLanguage;

        card.style.display = matchesSearch && matchesLanguage ? 'block' : 'none';
    });
}

function getFileLanguage(fileName) {
    const extension = fileName.split('.').pop().toLowerCase();
    const languages = {
        'js': 'javascript',
        'py': 'python',
        'java': 'java',
        'html': 'html',
        'css': 'css'
    };
    return languages[extension] || 'other';
}

async function viewCode(codeId) {
    showLoading(true);
    try {
        const doc = await db.collection('codes').doc(codeId).get();
        if (!doc.exists) {
            alert('代码文件不存在');
            return;
        }

        const code = doc.data();
        const modalTitle = document.getElementById('modalTitle');
        const fileName = document.getElementById('fileName');
        const codeContent = document.getElementById('codeContent');
        const codeForm = document.getElementById('codeForm');
        const modal = document.getElementById('codeModal');

        if (modalTitle && fileName && codeContent && codeForm && modal) {
            modalTitle.textContent = '查看代码';
            fileName.value = code.fileName;
            fileName.readOnly = true;
            codeContent.value = code.content;
            codeContent.readOnly = true;
            codeForm.onsubmit = null;
            modal.style.display = 'block';
        }

    } catch (error) {
        console.error('查看代码失败:', error);
        alert('操作失败: ' + error.message);
    } finally {
        showLoading(false);
    }
}

async function editCode(codeId) {
    showLoading(true);
    try {
        const doc = await db.collection('codes').doc(codeId).get();
        if (!doc.exists) {
            alert('代码文件不存在');
            return;
        }

        const code = doc.data();
        const modalTitle = document.getElementById('modalTitle');
        const codeIdInput = document.getElementById('codeId');
        const fileName = document.getElementById('fileName');
        const codeContent = document.getElementById('codeContent');
        const codeForm = document.getElementById('codeForm');
        const modal = document.getElementById('codeModal');

        if (modalTitle && codeIdInput && fileName && codeContent && codeForm && modal) {
            modalTitle.textContent = '编辑代码';
            codeIdInput.value = codeId;
            fileName.value = code.fileName;
            fileName.readOnly = false;
            codeContent.value = code.content;
            codeContent.readOnly = false;
            codeForm.onsubmit = handleCodeSubmit;
            modal.style.display = 'block';
        }

    } catch (error) {
        console.error('编辑代码失败:', error);
        alert('操作失败: ' + error.message);
    } finally {
        showLoading(false);
    }
}

async function deleteCode(codeId) {
    if (!confirm('确定要删除这个代码文件吗？此操作不可恢复。')) {
        return;
    }

    showLoading(true);
    try {
        await db.collection('codes').doc(codeId).delete();
        alert('代码文件已删除');
        if (currentUser.role === 'admin') {
            await loadAdminCodeList();
        } else {
            await loadUserCodeList();
        }
    } catch (error) {
        console.error('删除代码失败:', error);
        alert('操作失败: ' + error.message);
    } finally {
        showLoading(false);
    }
}

// Excel文件处理功能
function initializeExcelUpload() {
    const fileInput = document.getElementById('excelFileInput');
    const adminFileInput = document.getElementById('adminExcelFileInput');
    
    if (fileInput) {
        fileInput.addEventListener('change', handleExcelFileSelect);
    }
    
    if (adminFileInput) {
        adminFileInput.addEventListener('change', handleAdminExcelFileSelect);
    }
}

function handleExcelFileSelect(event) {
    handleExcelFileSelectCommon(event, false);
}

function handleAdminExcelFileSelect(event) {
    handleExcelFileSelectCommon(event, true);
}

function handleExcelFileSelectCommon(event, isAdmin) {
    const file = event.target.files[0];
    if (!file) return;

    // 验证文件类型
    const allowedTypes = [
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'application/vnd.ms-excel'
    ];
    
    if (!allowedTypes.includes(file.type)) {
        alert('请选择有效的Excel文件 (.xlsx 或 .xls)');
        return;
    }

    // 验证文件大小 (最大5MB)
    if (file.size > 5 * 1024 * 1024) {
        alert('文件大小不能超过5MB');
        return;
    }

    readExcelFile(file, isAdmin);
}

function readExcelFile(file, isAdmin = false) {
    const reader = new FileReader();
    
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            // 获取第一个工作表
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            
            // 转换为JSON数据
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            
            if (jsonData.length === 0) {
                alert('Excel文件为空');
                return;
            }

            currentExcelData = {
                fileName: file.name,
                data: jsonData,
                headers: jsonData[0] || [],
                rows: jsonData.slice(1),
                isAdmin: isAdmin
            };

            displayExcelPreview(isAdmin);
            
        } catch (error) {
            console.error('Excel文件读取失败:', error);
            alert('Excel文件读取失败，请检查文件格式');
        }
    };

    reader.readAsArrayBuffer(file);
}

function displayExcelPreview(isAdmin = false) {
    const previewId = isAdmin ? 'adminExcelPreview' : 'excelPreview';
    const containerId = isAdmin ? 'adminExcelTableContainer' : 'excelTableContainer';
    
    const previewDiv = document.getElementById(previewId);
    const tableContainer = document.getElementById(containerId);
    
    if (!previewDiv || !tableContainer || !currentExcelData) return;

    // 创建预览表格
    let tableHTML = `
        <div class="file-info">
            <p><strong>文件名:</strong> ${currentExcelData.fileName}</p>
            <p><strong>数据行数:</strong> ${currentExcelData.rows.length}</p>
            <p><strong>列数:</strong> ${currentExcelData.headers.length}</p>
        </div>
        <table class="excel-table">
            <thead>
                <tr>
                    ${currentExcelData.headers.map(header => 
                        `<th>${header || '未命名列'}</th>`
                    ).join('')}
                </tr>
            </thead>
            <tbody>
    `;

    // 只显示前10行数据作为预览
    const previewRows = currentExcelData.rows.slice(0, 10);
    previewRows.forEach(row => {
        tableHTML += '<tr>';
        currentExcelData.headers.forEach((_, index) => {
            tableHTML += `<td>${row[index] || ''}</td>`;
        });
        tableHTML += '</tr>';
    });

    if (currentExcelData.rows.length > 10) {
        tableHTML += `
            <tr>
                <td colspan="${currentExcelData.headers.length}" style="text-align: center; color: #666; font-style: italic;">
                    ... 还有 ${currentExcelData.rows.length - 10} 行数据
                </td>
            </tr>
        `;
    }

    tableHTML += '</tbody></table>';
    
    tableContainer.innerHTML = tableHTML;
    previewDiv.style.display = 'block';
}

async function processExcelFile() {
    await processExcelFileCommon(false);
}

async function processAdminExcelFile() {
    await processExcelFileCommon(true);
}

async function processExcelFileCommon(isAdmin) {
    if (!currentExcelData) {
        alert('请先选择Excel文件');
        return;
    }

    showLoading(true);
    try {
        // 将Excel数据保存到Firestore
        const excelDoc = {
            fileName: currentExcelData.fileName,
            headers: currentExcelData.headers,
            data: currentExcelData.rows,
            totalRows: currentExcelData.rows.length,
            uploadedAt: CURRENT_TIME,
            uploadedBy: currentUser.username,
            fileType: 'excel'
        };

        const docRef = await db.collection('excel_files').add(excelDoc);
        
        alert(`Excel文件上传成功！\n文件ID: ${docRef.id}\n数据行数: ${currentExcelData.rows.length}`);
        
        if (isAdmin) {
            clearAdminExcelPreview();
            await loadAdminExcelList();
        } else {
            clearExcelPreview();
            await loadUserExcelList();
        }

    } catch (error) {
        console.error('Excel文件上传失败:', error);
        alert('上传失败: ' + error.message);
    } finally {
        showLoading(false);
    }
}

function clearExcelPreview() {
    currentExcelData = null;
    const previewDiv = document.getElementById('excelPreview');
    const fileInput = document.getElementById('excelFileInput');
    
    if (previewDiv) previewDiv.style.display = 'none';
    if (fileInput) fileInput.value = '';
}

function clearAdminExcelPreview() {
    currentExcelData = null;
    const previewDiv = document.getElementById('adminExcelPreview');
    const fileInput = document.getElementById('adminExcelFileInput');
    
    if (previewDiv) previewDiv.style.display = 'none';
    if (fileInput) fileInput.value = '';
}

async function loadUserExcelList() {
    showLoading(true);
    try {
        const snapshot = await db.collection('excel_files')
            .where('uploadedBy', '==', currentUser.username)
            .orderBy('uploadedAt', 'desc')
            .get();

        displayExcelFileList(snapshot, 'userExcelList', false);
        
    } catch (error) {
        console.error('加载Excel文件列表失败:', error);
    } finally {
        showLoading(false);
    }
}

async function loadAdminExcelList() {
    showLoading(true);
    try {
        const snapshot = await db.collection('excel_files')
            .orderBy('uploadedAt', 'desc')
            .get();

        displayExcelFileList(snapshot, 'adminExcelList', true);
        
    } catch (error) {
        console.error('加载Excel文件列表失败:', error);
    } finally {
        showLoading(false);
    }
}

function displayExcelFileList(snapshot, containerId, isAdmin) {
    const container = document.getElementById(containerId);
    if (!container) return;

    if (snapshot.empty) {
        container.innerHTML = `
            <div class="empty-state">
                <i class="ri-file-excel-2-line" style="font-size: 48px; color: #666;"></i>
                <p>暂无Excel文件</p>
            </div>
        `;
        return;
    }

    const filesHTML = snapshot.docs.map(doc => {
        const file = doc.data();
        return `
            <div class="file-card">
                <div class="file-header">
                    <i class="ri-file-excel-2-line" style="color: #28a745; font-size: 24px;"></i>
                    <span class="file-name">${file.fileName}</span>
                </div>
                <div class="file-details">
                    <p>上传者: ${file.uploadedBy}</p>
                    <p>数据行数: ${file.totalRows}</p>
                    <p>上传时间: ${formatDate(file.uploadedAt)}</p>
                </div>
                <div class="file-actions">
                    <button onclick="viewExcelFile('${doc.id}')" class="btn btn-sm btn-primary">查看</button>
                    <button onclick="downloadExcelData('${doc.id}')" class="btn btn-sm btn-secondary">下载</button>
                    ${isAdmin || file.uploadedBy === currentUser.username ? 
                        `<button onclick="deleteExcelFile('${doc.id}')" class="btn btn-sm btn-danger">删除</button>` : ''}
                </div>
            </div>
        `;
    }).join('');

    container.innerHTML = filesHTML;
}

async function viewExcelFile(fileId) {
    try {
        const doc = await db.collection('excel_files').doc(fileId).get();
        if (!doc.exists) {
            alert('文件不存在');
            return;
        }

        const fileData = doc.data();
        
        // 创建查看窗口
        const modal = document.createElement('div');
        modal.className = 'modal';
        modal.style.display = 'block';
        modal.innerHTML = `
            <div class="modal-content excel-view-modal">
                <div class="modal-header">
                    <h3>${fileData.fileName}</h3>
                    <button onclick="this.closest('.modal').remove()" class="modal-close">&times;</button>
                </div>
                <div class="modal-body">
                    <div class="file-info">
                        <p><strong>上传者:</strong> ${fileData.uploadedBy}</p>
                        <p><strong>上传时间:</strong> ${formatDate(fileData.uploadedAt)}</p>
                        <p><strong>数据行数:</strong> ${fileData.totalRows}</p>
                    </div>
                    <div class="excel-data-container">
                        <table class="excel-table">
                            <thead>
                                <tr>
                                    ${fileData.headers.map(header => 
                                        `<th>${header || '未命名列'}</th>`
                                    ).join('')}
                                </tr>
                            </thead>
                            <tbody>
                                ${fileData.data.slice(0, 50).map(row => `
                                    <tr>
                                        ${fileData.headers.map((_, index) => 
                                            `<td>${row[index] || ''}</td>`
                                        ).join('')}
                                    </tr>
                                `).join('')}
                                ${fileData.data.length > 50 ? 
                                    `<tr><td colspan="${fileData.headers.length}" style="text-align: center; color: #666;">... 还有 ${fileData.data.length - 50} 行数据</td></tr>` : ''}
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        `;

        document.body.appendChild(modal);

    } catch (error) {
        console.error('查看文件失败:', error);
        alert('查看文件失败: ' + error.message);
    }
}

async function downloadExcelData(fileId) {
    try {
        const doc = await db.collection('excel_files').doc(fileId).get();
        if (!doc.exists) {
            alert('文件不存在');
            return;
        }

        const fileData = doc.data();
        
        // 创建新的工作簿
        const wb = XLSX.utils.book_new();
        const wsData = [fileData.headers, ...fileData.data];
        const ws = XLSX.utils.aoa_to_sheet(wsData);
        
        XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
        
        // 下载文件
        XLSX.writeFile(wb, fileData.fileName);

    } catch (error) {
        console.error('下载失败:', error);
        alert('下载失败: ' + error.message);
    }
}

async function deleteExcelFile(fileId) {
    if (!confirm('确定要删除这个Excel文件吗？')) return;

    showLoading(true);
    try {
        await db.collection('excel_files').doc(fileId).delete();
        alert('文件已删除');
        
        // 刷新列表
        if (currentUser.role === 'admin') {
            await loadAdminExcelList();
        } else {
            await loadUserExcelList();
        }

    } catch (error) {
        console.error('删除失败:', error);
        alert('删除失败: ' + error.message);
    } finally {
        showLoading(false);
    }
}

function handleLogout() {
    currentUser = null;
    showLoginForm();
}

// 标签页切换功能
function switchTab(tabId) {
    // 隐藏所有管理员标签内容
    const adminTabs = ['pendingUsers', 'codeFiles', 'excelFiles'];
    adminTabs.forEach(id => {
        const element = document.getElementById(id);
        if (element) element.classList.remove('active');
    });
    
    // 移除所有管理员标签按钮的active类
    document.querySelectorAll('#adminPanel .tab-btn').forEach(btn => {
        btn.classList.remove('active');
    });
    
    // 显示选中的标签内容
    const selectedTab = document.getElementById(tabId);
    if (selectedTab) {
        selectedTab.classList.add('active');
    }
    
    // 添加active类到对应按钮
    const selectedBtn = document.querySelector(`#adminPanel [onclick="switchTab('${tabId}')"]`);
    if (selectedBtn) {
        selectedBtn.classList.add('active');
    }
    
    // 根据标签加载相应数据
    if (tabId === 'codeFiles') {
        loadAdminCodeList();
    } else if (tabId === 'pendingUsers') {
        loadPendingUsers();
    } else if (tabId === 'excelFiles') {
        loadAdminExcelList();
    }
}

function switchUserTab(tabId) {
    // 隐藏所有用户标签内容
    const userTabs = ['codeTab', 'excelTab'];
    userTabs.forEach(id => {
        const element = document.getElementById(id);
        if (element) element.classList.remove('active');
    });
    
    // 移除所有用户标签按钮的active类
    document.querySelectorAll('#userPanel .tab-btn').forEach(btn => {
        btn.classList.remove('active');
    });
    
    // 显示选中的标签内容
    const selectedTab = document.getElementById(tabId);
    if (selectedTab) {
        selectedTab.classList.add('active');
    }
    
    // 添加active类到对应按钮
    const selectedBtn = document.querySelector(`#userPanel [onclick="switchUserTab('${tabId}')"]`);
    if (selectedBtn) {
        selectedBtn.classList.add('active');
    }
    
    // 根据标签加载相应数据
    if (tabId === 'codeTab') {
        loadUserCodeList();
    } else if (tabId === 'excelTab') {
        loadUserExcelList();
    }
}

// 页面加载完成后初始化
document.addEventListener('DOMContentLoaded', () => {
    console.log('系统初始化 -', CURRENT_TIME);
    console.log('当前用户:', CURRENT_USER);

    // 绑定表单提交事件
    const loginForm = document.getElementById('loginFormElement');
    const registerForm = document.getElementById('registerFormElement');
    const statusQueryForm = document.getElementById('statusQueryElement');
    const codeForm = document.getElementById('codeForm');

    if (loginForm) loginForm.addEventListener('submit', handleLogin);
    if (registerForm) registerForm.addEventListener('submit', handleRegister);
    if (statusQueryForm) statusQueryForm.addEventListener('submit', handleStatusQuery);
    if (codeForm) codeForm.addEventListener('submit', handleCodeSubmit);

    // 初始化Excel上传功能
    initializeExcelUpload();

    // 显示登录表单
    showLoginForm();
});

// 导出全局函数
window.showLoginForm = showLoginForm;
window.showRegisterForm = showRegisterForm;
window.showStatusQuery = showStatusQuery;
window.handleLogout = handleLogout;
window.showUploadForm = showUploadForm;
window.hideCodeModal = hideCodeModal;
window.handleApprove = handleApprove;
window.handleReject = handleReject;
window.viewCode = viewCode;
window.editCode = editCode;
window.deleteCode = deleteCode;
window.switchTab = switchTab;
window.switchUserTab = switchUserTab;
window.processExcelFile = processExcelFile;
window.processAdminExcelFile = processAdminExcelFile;
window.clearExcelPreview = clearExcelPreview;
window.clearAdminExcelPreview = clearAdminExcelPreview;
window.viewExcelFile = viewExcelFile;
window.downloadExcelData = downloadExcelData;
window.deleteExcelFile = deleteExcelFile;
