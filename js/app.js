// 全局常量
const CURRENT_TIME = '2025-07-30 09:28:13';
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

// 完全扁平化的数据处理函数 - 彻底解决嵌套数组问题
function totallyFlattenData(data) {
    console.log('开始完全扁平化处理:', typeof data, data);
    
    // 处理 null 和 undefined
    if (data === null || data === undefined) {
        return '';
    }
    
    // 处理基本类型
    if (typeof data === 'string') {
        return data;
    }
    
    if (typeof data === 'number') {
        if (isNaN(data) || !isFinite(data)) {
            return '';
        }
        return data;
    }
    
    if (typeof data === 'boolean') {
        return data.toString();
    }
    
    // 处理日期
    if (data instanceof Date) {
        return data.toISOString();
    }
    
    // 关键：将任何数组或对象都转换为字符串
    if (Array.isArray(data) || (typeof data === 'object' && data !== null)) {
        try {
            return JSON.stringify(data);
        } catch (error) {
            console.warn('JSON序列化失败，转换为字符串:', error);
            return String(data);
        }
    }
    
    // 其他情况转换为字符串
    return String(data);
}

// 改进的Excel数据处理函数 - 确保数据完全扁平化
function processExcelDataForFirestore(excelData) {
    console.log('开始处理Excel数据，原始数据结构:', excelData);
    
    // 强化headers处理 - 确保每个header都是字符串
    const sanitizedHeaders = [];
    if (Array.isArray(excelData.headers)) {
        for (let i = 0; i < excelData.headers.length; i++) {
            const header = excelData.headers[i];
            let cleanHeader = '';
            
            if (header === null || header === undefined || header === '') {
                cleanHeader = `列${i + 1}`;
            } else if (Array.isArray(header)) {
                // 如果header本身是数组，转换为字符串
                cleanHeader = `列${i + 1}_${JSON.stringify(header)}`;
            } else if (typeof header === 'object' && header !== null) {
                // 如果header是对象，转换为字符串
                cleanHeader = `列${i + 1}_${JSON.stringify(header)}`;
            } else {
                cleanHeader = String(header);
            }
            
            sanitizedHeaders.push(cleanHeader);
        }
    }
    
    // 彻底扁平化行数据处理
    const sanitizedRows = [];
    if (Array.isArray(excelData.rows)) {
        for (let rowIndex = 0; rowIndex < excelData.rows.length; rowIndex++) {
            const row = excelData.rows[rowIndex];
            const sanitizedRow = {};
            
            if (Array.isArray(row)) {
                // 将行数据转换为对象，使用索引作为key
                for (let colIndex = 0; colIndex < sanitizedHeaders.length; colIndex++) {
                    const headerKey = `col_${colIndex}`;
                    const cell = row[colIndex];
                    const flattenedCell = totallyFlattenData(cell);
                    
                    // 二次验证：确保flattenedCell不是数组或对象
                    let finalValue = flattenedCell;
                    if (Array.isArray(finalValue) || (typeof finalValue === 'object' && finalValue !== null)) {
                        finalValue = JSON.stringify(finalValue);
                    }
                    
                    sanitizedRow[headerKey] = finalValue;
                    console.log(`处理单元格[${rowIndex}][${colIndex}]:`, typeof cell, '=>', typeof finalValue, finalValue);
                }
            } else {
                // 如果行不是数组，创建空对象
                for (let colIndex = 0; colIndex < sanitizedHeaders.length; colIndex++) {
                    const headerKey = `col_${colIndex}`;
                    sanitizedRow[headerKey] = '';
                }
            }
            
            sanitizedRows.push(sanitizedRow);
        }
    }
    
    const processedData = {
        fileName: String(excelData.fileName || '未知文件'),
        headers: sanitizedHeaders,
        rows: sanitizedRows,
        totalRows: sanitizedRows.length
    };
    
    console.log('处理后的数据结构验证:', {
        fileName: typeof processedData.fileName,
        headersType: typeof processedData.headers,
        headersLength: processedData.headers.length,
        headersIsArray: Array.isArray(processedData.headers),
        rowsType: typeof processedData.rows,
        rowsLength: processedData.rows.length,
        rowsIsArray: Array.isArray(processedData.rows),
        firstRowType: processedData.rows[0] ? typeof processedData.rows[0] : 'undefined',
        firstRowIsObject: processedData.rows[0] ? (typeof processedData.rows[0] === 'object' && !Array.isArray(processedData.rows[0])) : false,
        sampleRowStructure: processedData.rows[0] ? Object.keys(processedData.rows[0]).slice(0, 3) : []
    });
    
    return processedData;
}

// 更严格的数据验证函数
function strictValidateFirestoreData(data, path = '') {
    console.log(`验证路径 ${path}:`, typeof data, Array.isArray(data) ? '数组' : '非数组');
    
    if (Array.isArray(data)) {
        // 验证数组中的每个元素
        for (let i = 0; i < data.length; i++) {
            const item = data[i];
            const currentPath = `${path}[${i}]`;
            
            // 严格检查：数组元素不能是数组或复杂对象
            if (Array.isArray(item)) {
                console.error(`❌ 发现嵌套数组在路径: ${currentPath}`, item);
                return false;
            }
            
            // 检查对象类型
            if (typeof item === 'object' && item !== null) {
                // 只允许简单对象，不允许嵌套数组
                if (!strictValidateFirestoreData(item, currentPath)) {
                    return false;
                }
            }
        }
    } else if (typeof data === 'object' && data !== null) {
        // 验证对象的每个属性
        for (const [key, value] of Object.entries(data)) {
            const currentPath = path ? `${path}.${key}` : key;
            
            // 检查属性值类型
            if (Array.isArray(value)) {
                // 数组值必须只包含基本类型
                for (let i = 0; i < value.length; i++) {
                    const arrayItem = value[i];
                    const arrayPath = `${currentPath}[${i}]`;
                    
                    if (Array.isArray(arrayItem)) {
                        console.error(`❌ 发现嵌套数组在路径: ${arrayPath}`, arrayItem);
                        return false;
                    }
                    
                    if (typeof arrayItem === 'object' && arrayItem !== null) {
                        console.error(`❌ 发现数组中包含对象在路径: ${arrayPath}`, arrayItem);
                        return false;
                    }
                }
            } else if (typeof value === 'object' && value !== null) {
                // 递归验证嵌套对象
                if (!strictValidateFirestoreData(value, currentPath)) {
                    return false;
                }
            }
        }
    }
    
    console.log(`✅ 路径 ${path} 验证通过`);
    return true;
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
                        `<th>${String(header || '未命名列')}</th>`
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
            const cellValue = row[index];
            let displayValue = '';
            
            if (Array.isArray(cellValue)) {
                displayValue = JSON.stringify(cellValue);
            } else if (cellValue && typeof cellValue === 'object') {
                displayValue = JSON.stringify(cellValue);
            } else {
                displayValue = String(cellValue || '');
            }
            
            tableHTML += `<td>${displayValue}</td>`;
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
        console.log('🚀 开始处理Excel数据...');
        console.log('📊 原始数据:', currentExcelData);
        
        // 使用完全扁平化的Excel数据处理函数
        const processedData = processExcelDataForFirestore(currentExcelData);
        
        console.log('📋 处理后数据概要:', processedData);
        
        // 创建最终的文档对象 - 使用Map替代数组来存储headers
        const excelDoc = {
            fileName: processedData.fileName,
            // 将headers转换为对象而不是数组，避免任何数组嵌套
            headersMap: {},
            headersList: processedData.headers.join('|||'), // 使用字符串存储headers
            data: processedData.rows,
            totalRows: processedData.totalRows,
            uploadedAt: CURRENT_TIME,
            uploadedBy: currentUser.username,
            fileType: 'excel'
        };
        
        // 创建headers映射
        for (let i = 0; i < processedData.headers.length; i++) {
            excelDoc.headersMap[`header_${i}`] = processedData.headers[i];
        }

        console.log('📝 最终保存的文档结构:', {
            fileName: `"${excelDoc.fileName}" (${typeof excelDoc.fileName})`,
            headersMapType: typeof excelDoc.headersMap,
            headersListType: typeof excelDoc.headersList,
            dataRowsCount: `${excelDoc.data.length} (${typeof excelDoc.data})`,
            totalRows: `${excelDoc.totalRows} (${typeof excelDoc.totalRows})`,
            uploadedAt: `"${excelDoc.uploadedAt}" (${typeof excelDoc.uploadedAt})`,
            uploadedBy: `"${excelDoc.uploadedBy}" (${typeof excelDoc.uploadedBy})`,
            fileType: `"${excelDoc.fileType}" (${typeof excelDoc.fileType})`
        });

        // 严格验证处理后的数据
        console.log('🔍 开始严格验证数据结构...');
        if (!strictValidateFirestoreData(excelDoc)) {
            throw new Error('❌ 数据验证失败：仍然包含不支持的嵌套结构');
        }
        
        console.log('✅ 数据验证通过，准备保存到Firestore');

        console.log('💾 开始保存到Firestore...');
        const docRef = await db.collection('excel_files').add(excelDoc);
        
        console.log('🎉 Excel文件上传成功！文档ID:', docRef.id);
        alert(`Excel文件上传成功！\n文件ID: ${docRef.id}\n数据行数: ${processedData.totalRows}`);
        
        if (isAdmin) {
            clearAdminExcelPreview();
            await loadAdminExcelList();
        } else {
            clearExcelPreview();
            await loadUserExcelList();
        }

    } catch (error) {
        console.error('❌ Excel文件上传失败:', error);
        console.error('📋 错误详情:', error.message);
        console.error('📚 错误堆栈:', error.stack);
        alert('上传失败: ' + error.message);
    } finally {
        showLoading(false);
    }
}

// 更新查看文件函数以适应新的数据结构
async function viewExcelFile(fileId) {
    try {
        const doc = await db.collection('excel_files').doc(fileId).get();
        if (!doc.exists) {
            alert('文件不存在');
            return;
        }

        const fileData = doc.data();
        
        // 处理headers - 支持新旧格式
        let headers = [];
        if (fileData.headersList) {
            // 新格式：从字符串恢复headers
            headers = fileData.headersList.split('|||');
        } else if (fileData.headers) {
            // 旧格式：直接使用headers数组
            headers = fileData.headers;
        } else if (fileData.headersMap) {
            // 从headersMap恢复headers
            const headerKeys = Object.keys(fileData.headersMap).sort((a, b) => {
                const aIndex = parseInt(a.split('_')[1]);
                const bIndex = parseInt(b.split('_')[1]);
                return aIndex - bIndex;
            });
            headers = headerKeys.map(key => fileData.headersMap[key]);
        }
        
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
                                    ${headers.map(header => 
                                        `<th>${String(header || '未命名列')}</th>`
                                    ).join('')}
                                </tr>
                            </thead>
                            <tbody>
                                ${fileData.data.slice(0, 50).map(row => `
                                    <tr>
                                        ${headers.map((_, index) => {
                                            const colKey = `col_${index}`;
                                            const cellValue = row[colKey];
                                            let displayValue = '';
                                            
                                            if (typeof cellValue === 'string') {
                                                displayValue = cellValue;
                                            } else if (cellValue === null || cellValue === undefined) {
                                                displayValue = '';
                                            } else {
                                                displayValue = String(cellValue);
                                            }
                                            
                                            return `<td>${displayValue}</td>`;
                                        }).join('')}
                                    </tr>
                                `).join('')}
                                ${fileData.data.length > 50 ? 
                                    `<tr><td colspan="${headers.length}" style="text-align: center; color: #666;">... 还有 ${fileData.data.length - 50} 行数据</td></tr>` : ''}
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

// 更新下载函数以适应新的数据结构
async function downloadExcelData(fileId) {
    try {
        const doc = await db.collection('excel_files').doc(fileId).get();
        if (!doc.exists) {
            alert('文件不存在');
            return;
        }

        const fileData = doc.data();
        
        // 处理headers - 支持新旧格式
        let headers = [];
        if (fileData.headersList) {
            headers = fileData.headersList.split('|||');
        } else if (fileData.headers) {
            headers = fileData.headers;
        } else if (fileData.headersMap) {
            const headerKeys = Object.keys(fileData.headersMap).sort((a, b) => {
                const aIndex = parseInt(a.split('_')[1]);
                const bIndex = parseInt(b.split('_')[1]);
                return aIndex - bIndex;
            });
            headers = headerKeys.map(key => fileData.headersMap[key]);
        }
        
        // 将对象数组转换回二维数组格式
        const wsData = [headers];
        fileData.data.forEach(row => {
            const rowArray = [];
            headers.forEach((_, index) => {
                const colKey = `col_${index}`;
                rowArray.push(row[colKey] || '');
            });
            wsData.push(rowArray);
        });
        
        // 创建新的工作簿
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet(wsData);
        
        XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
        
        // 下载文件
        XLSX.writeFile(wb, fileData.fileName);

    } catch (error) {
        console.error('下载失败:', error);
        alert('下载失败: ' + error.message);
    }
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
