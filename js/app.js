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
    
    if (excelData.isMultiSheet) {
        // 处理多工作表数据
        const processedData = {
            fileName: String(excelData.fileName || '未知文件'),
            fileSize: excelData.fileSize,
            fileType: excelData.fileType,
            lastModified: excelData.lastModified,
            totalSheets: excelData.totalSheets,
            sheetNames: excelData.sheetNames.join('|||'),
            isMultiSheet: true,
            sheets: {},
            totalRows: 0
        };
        
        // 处理每个工作表
        excelData.sheets.forEach((sheet, sheetIndex) => {
            const sheetKey = `sheet_${sheetIndex}`;
            const sanitizedHeaders = [];
            
            // 处理headers
            for (let i = 0; i < sheet.headers.length; i++) {
                const header = sheet.headers[i];
                let cleanHeader = String(header || `列${i + 1}`);
                sanitizedHeaders.push(cleanHeader);
            }
            
            // 处理行数据
            const dataRows = {};
            sheet.rows.forEach((row, rowIndex) => {
                const sanitizedRow = {};
                for (let colIndex = 0; colIndex < sanitizedHeaders.length; colIndex++) {
                    const headerKey = `col_${colIndex}`;
                    const cell = row[colIndex];
                    const finalValue = totallyFlattenData(cell);
                    sanitizedRow[headerKey] = finalValue;
                }
                dataRows[`row_${rowIndex}`] = sanitizedRow;
            });
            
            processedData.sheets[sheetKey] = {
                sheetName: sheet.sheetName,
                headersList: sanitizedHeaders.join('|||'),
                headersMap: {},
                dataRows: dataRows,
                totalRows: sheet.rows.length
            };
            
            // 创建headers映射
            for (let i = 0; i < sanitizedHeaders.length; i++) {
                processedData.sheets[sheetKey].headersMap[`header_${i}`] = sanitizedHeaders[i];
            }
            
            processedData.totalRows += sheet.rows.length;
        });
        
        return processedData;
    } else {
        // 处理单工作表数据（原有逻辑）
        const sanitizedHeaders = [];
        if (Array.isArray(excelData.headers)) {
            for (let i = 0; i < excelData.headers.length; i++) {
                const header = excelData.headers[i];
                let cleanHeader = String(header || `列${i + 1}`);
                sanitizedHeaders.push(cleanHeader);
            }
        }
        
        const sanitizedRows = [];
        if (Array.isArray(excelData.rows)) {
            for (let rowIndex = 0; rowIndex < excelData.rows.length; rowIndex++) {
                const row = excelData.rows[rowIndex];
                const sanitizedRow = {};
                
                if (Array.isArray(row)) {
                    for (let colIndex = 0; colIndex < sanitizedHeaders.length; colIndex++) {
                        const headerKey = `col_${colIndex}`;
                        const cell = row[colIndex];
                        const finalValue = totallyFlattenData(cell);
                        sanitizedRow[headerKey] = finalValue;
                    }
                }
                
                sanitizedRows.push(sanitizedRow);
            }
        }
        
        const processedData = {
            fileName: String(excelData.fileName || '未知文件'),
            fileSize: excelData.fileSize,
            fileType: excelData.fileType,
            lastModified: excelData.lastModified,
            sheetName: excelData.sheetName,
            totalSheets: excelData.totalSheets,
            sheetNames: excelData.sheetNames.join('|||'),
            isMultiSheet: false,
            headersList: sanitizedHeaders.join('|||'),
            headersMap: {},
            dataRows: {},
            totalRows: sanitizedRows.length
        };
        
        // 创建headers映射
        for (let i = 0; i < sanitizedHeaders.length; i++) {
            processedData.headersMap[`header_${i}`] = sanitizedHeaders[i];
        }
        
        // 创建数据行映射
        for (let rowIndex = 0; rowIndex < sanitizedRows.length; rowIndex++) {
            processedData.dataRows[`row_${rowIndex}`] = sanitizedRows[rowIndex];
        }
        
        return processedData;
    }
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

// 显示工作表选择器
function showSheetSelector(workbook, file, isAdmin) {
    const modal = document.createElement('div');
    modal.className = 'modal';
    modal.style.display = 'block';
    modal.innerHTML = `
        <div class="modal-content">
            <div class="modal-header">
                <h3>选择工作表</h3>
                <button onclick="this.closest('.modal').remove()" class="modal-close">&times;</button>
            </div>
            <div class="modal-body">
                <p>检测到多个工作表，请选择要上传的工作表：</p>
                <div class="sheet-selector">
                    ${workbook.SheetNames.map((sheetName, index) => {
                        const worksheet = workbook.Sheets[sheetName];
                        const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
                        const rowCount = range.e.r - range.s.r;
                        const colCount = range.e.c - range.s.c + 1;
                        
                        return `
                            <div class="sheet-option" onclick="selectSheet('${sheetName}', ${index})">
                                <div class="sheet-info">
                                    <h4>${sheetName}</h4>
                                    <p>行数: ${rowCount}, 列数: ${colCount}</p>
                                </div>
                                <i class="ri-arrow-right-line"></i>
                            </div>
                        `;
                    }).join('')}
                </div>
                <div class="modal-actions">
                    <button onclick="selectAllSheets()" class="btn btn-primary">
                        <i class="ri-stack-line"></i> 全部上传
                    </button>
                    <button onclick="this.closest('.modal').remove()" class="btn btn-secondary">取消</button>
                </div>
            </div>
        </div>
    `;

    document.body.appendChild(modal);
    
    // 保存工作簿信息到全局变量
    window.currentWorkbook = workbook;
    window.currentFile = file;
    window.currentIsAdmin = isAdmin;
}

// 选择单个工作表
function selectSheet(sheetName, index) {
    const workbook = window.currentWorkbook;
    const file = window.currentFile;
    const isAdmin = window.currentIsAdmin;
    
    processWorksheet(workbook, sheetName, file, isAdmin);
    
    // 关闭选择器
    document.querySelector('.modal').remove();
}

// 选择所有工作表
function selectAllSheets() {
    const workbook = window.currentWorkbook;
    const file = window.currentFile;
    const isAdmin = window.currentIsAdmin;
    
    // 处理所有工作表
    const allSheetsData = [];
    
    workbook.SheetNames.forEach(sheetName => {
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
            header: 1,
            defval: '', 
            raw: false  
        });
        
        if (jsonData.length > 0) {
            const cleanHeaders = jsonData[0] ? jsonData[0].map((header, index) => {
                if (header === null || header === undefined || header === '') {
                    return `列${index + 1}`;
                }
                return String(header);
            }) : [];

            const cleanRows = jsonData.slice(1).map(row => {
                if (!Array.isArray(row)) {
                    return new Array(cleanHeaders.length).fill('');
                }
                return row.map(cell => {
                    if (cell === null || cell === undefined) {
                        return '';
                    }
                    if (typeof cell === 'object') {
                        return JSON.stringify(cell);
                    }
                    return String(cell);
                });
            });

            allSheetsData.push({
                sheetName: sheetName,
                headers: cleanHeaders,
                rows: cleanRows
            });
        }
    });

    currentExcelData = {
        fileName: file.name,
        originalFile: file, // 保存原始文件
        fileSize: file.size,
        fileType: file.type,
        lastModified: file.lastModified,
        totalSheets: workbook.SheetNames.length,
        sheetNames: workbook.SheetNames,
        sheets: allSheetsData,
        isMultiSheet: true,
        isAdmin: isAdmin
    };

    console.log('✅ 多工作表Excel数据:', {
        fileName: currentExcelData.fileName,
        totalSheets: currentExcelData.totalSheets,
        sheetNames: currentExcelData.sheetNames,
        sheetsData: currentExcelData.sheets.map(sheet => ({
            name: sheet.sheetName,
            rows: sheet.rows.length,
            cols: sheet.headers.length
        }))
    });

    displayExcelPreview(isAdmin);
    
    // 关闭选择器
    document.querySelector('.modal').remove();
}

// 处理单个工作表
function processWorksheet(workbook, sheetName, file, isAdmin) {
    const worksheet = workbook.Sheets[sheetName];
    
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
        header: 1,
        defval: '', 
        raw: false  
    });
    
    if (jsonData.length === 0) {
        alert('选择的工作表为空');
        return;
    }

    const cleanHeaders = jsonData[0] ? jsonData[0].map((header, index) => {
        if (header === null || header === undefined || header === '') {
            return `列${index + 1}`;
        }
        return String(header);
    }) : [];

    const cleanRows = jsonData.slice(1).map(row => {
        if (!Array.isArray(row)) {
            return new Array(cleanHeaders.length).fill('');
        }
        return row.map(cell => {
            if (cell === null || cell === undefined) {
                return '';
            }
            if (typeof cell === 'object') {
                return JSON.stringify(cell);
            }
            return String(cell);
        });
    });

    currentExcelData = {
        fileName: file.name,
        originalFile: file, // 保存原始文件
        fileSize: file.size,
        fileType: file.type,
        lastModified: file.lastModified,
        sheetName: sheetName,
        totalSheets: workbook.SheetNames.length,
        sheetNames: workbook.SheetNames,
        data: [cleanHeaders, ...cleanRows],
        headers: cleanHeaders,
        rows: cleanRows,
        isMultiSheet: false,
        isAdmin: isAdmin
    };

    console.log('✅ 单工作表Excel数据:', {
        fileName: currentExcelData.fileName,
        sheetName: currentExcelData.sheetName,
        totalSheets: currentExcelData.totalSheets,
        headersCount: currentExcelData.headers.length,
        rowsCount: currentExcelData.rows.length
    });

    displayExcelPreview(isAdmin);
}

// 更新预览显示函数
function displayExcelPreview(isAdmin = false) {
    const previewId = isAdmin ? 'adminExcelPreview' : 'excelPreview';
    const containerId = isAdmin ? 'adminExcelTableContainer' : 'excelTableContainer';
    
    const previewDiv = document.getElementById(previewId);
    const tableContainer = document.getElementById(containerId);
    
    if (!previewDiv || !tableContainer || !currentExcelData) return;

    let tableHTML = `
        <div class="file-info">
            <p><strong>文件名:</strong> ${currentExcelData.fileName}</p>
            <p><strong>文件大小:</strong> ${(currentExcelData.fileSize / 1024).toFixed(2)} KB</p>
            <p><strong>总工作表数:</strong> ${currentExcelData.totalSheets}</p>
    `;

    if (currentExcelData.isMultiSheet) {
        // 多工作表预览
        tableHTML += `<p><strong>将要上传:</strong> 所有工作表 (${currentExcelData.sheets.length}个)</p></div>`;
        
        currentExcelData.sheets.forEach((sheet, index) => {
            if (index < 3) { // 只显示前3个工作表的预览
                tableHTML += `
                    <div class="sheet-preview">
                        <h4>工作表: ${sheet.sheetName}</h4>
                        <p>数据行数: ${sheet.rows.length}, 列数: ${sheet.headers.length}</p>
                        <table class="excel-table">
                            <thead>
                                <tr>
                                    ${sheet.headers.map(header => 
                                        `<th>${String(header || '未命名列')}</th>`
                                    ).join('')}
                                </tr>
                            </thead>
                            <tbody>
                                ${sheet.rows.slice(0, 5).map(row => `
                                    <tr>
                                        ${sheet.headers.map((_, colIndex) => {
                                            const cellValue = row[colIndex];
                                            let displayValue = String(cellValue || '');
                                            return `<td>${displayValue}</td>`;
                                        }).join('')}
                                    </tr>
                                `).join('')}
                                ${sheet.rows.length > 5 ? 
                                    `<tr><td colspan="${sheet.headers.length}" style="text-align: center; color: #666;">... 还有 ${sheet.rows.length - 5} 行</td></tr>` : ''}
                            </tbody>
                        </table>
                    </div>
                `;
            }
        });
        
        if (currentExcelData.sheets.length > 3) {
            tableHTML += `<p style="text-align: center; color: #666;">... 还有 ${currentExcelData.sheets.length - 3} 个工作表</p>`;
        }
    } else {
        // 单工作表预览
        tableHTML += `
            <p><strong>工作表:</strong> ${currentExcelData.sheetName}</p>
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
                ${currentExcelData.rows.slice(0, 10).map(row => `
                    <tr>
                        ${currentExcelData.headers.map((_, index) => {
                            const cellValue = row[index];
                            let displayValue = String(cellValue || '');
                            return `<td>${displayValue}</td>`;
                        }).join('')}
                    </tr>
                `).join('')}
                ${currentExcelData.rows.length > 10 ? 
                    `<tr><td colspan="${currentExcelData.headers.length}" style="text-align: center; color: #666;">... 还有 ${currentExcelData.rows.length - 10} 行数据</td></tr>` : ''}
            </tbody>
        </table>
        `;
    }
    
    tableContainer.innerHTML = tableHTML;
    previewDiv.style.display = 'block';
}

// 更新数据处理函数以支持多工作表
function processExcelDataForFirestore(excelData) {
    console.log('开始处理Excel数据，原始数据结构:', excelData);
    
    if (excelData.isMultiSheet) {
        // 处理多工作表数据
        const processedData = {
            fileName: String(excelData.fileName || '未知文件'),
            fileSize: excelData.fileSize,
            fileType: excelData.fileType,
            lastModified: excelData.lastModified,
            totalSheets: excelData.totalSheets,
            sheetNames: excelData.sheetNames.join('|||'),
            isMultiSheet: true,
            sheets: {},
            totalRows: 0
        };
        
        // 处理每个工作表
        excelData.sheets.forEach((sheet, sheetIndex) => {
            const sheetKey = `sheet_${sheetIndex}`;
            const sanitizedHeaders = [];
            
            // 处理headers
            for (let i = 0; i < sheet.headers.length; i++) {
                const header = sheet.headers[i];
                let cleanHeader = String(header || `列${i + 1}`);
                sanitizedHeaders.push(cleanHeader);
            }
            
            // 处理行数据
            const dataRows = {};
            sheet.rows.forEach((row, rowIndex) => {
                const sanitizedRow = {};
                for (let colIndex = 0; colIndex < sanitizedHeaders.length; colIndex++) {
                    const headerKey = `col_${colIndex}`;
                    const cell = row[colIndex];
                    const finalValue = totallyFlattenData(cell);
                    sanitizedRow[headerKey] = finalValue;
                }
                dataRows[`row_${rowIndex}`] = sanitizedRow;
            });
            
            processedData.sheets[sheetKey] = {
                sheetName: sheet.sheetName,
                headersList: sanitizedHeaders.join('|||'),
                headersMap: {},
                dataRows: dataRows,
                totalRows: sheet.rows.length
            };
            
            // 创建headers映射
            for (let i = 0; i < sanitizedHeaders.length; i++) {
                processedData.sheets[sheetKey].headersMap[`header_${i}`] = sanitizedHeaders[i];
            }
            
            processedData.totalRows += sheet.rows.length;
        });
        
        return processedData;
    } else {
        // 处理单工作表数据（原有逻辑）
        const sanitizedHeaders = [];
        if (Array.isArray(excelData.headers)) {
            for (let i = 0; i < excelData.headers.length; i++) {
                const header = excelData.headers[i];
                let cleanHeader = String(header || `列${i + 1}`);
                sanitizedHeaders.push(cleanHeader);
            }
        }
        
        const sanitizedRows = [];
        if (Array.isArray(excelData.rows)) {
            for (let rowIndex = 0; rowIndex < excelData.rows.length; rowIndex++) {
                const row = excelData.rows[rowIndex];
                const sanitizedRow = {};
                
                if (Array.isArray(row)) {
                    for (let colIndex = 0; colIndex < sanitizedHeaders.length; colIndex++) {
                        const headerKey = `col_${colIndex}`;
                        const cell = row[colIndex];
                        const finalValue = totallyFlattenData(cell);
                        sanitizedRow[headerKey] = finalValue;
                    }
                }
                
                sanitizedRows.push(sanitizedRow);
            }
        }
        
        const processedData = {
            fileName: String(excelData.fileName || '未知文件'),
            fileSize: excelData.fileSize,
            fileType: excelData.fileType,
            lastModified: excelData.lastModified,
            sheetName: excelData.sheetName,
            totalSheets: excelData.totalSheets,
            sheetNames: excelData.sheetNames.join('|||'),
            isMultiSheet: false,
            headersList: sanitizedHeaders.join('|||'),
            headersMap: {},
            dataRows: {},
            totalRows: sanitizedRows.length
        };
        
        // 创建headers映射
        for (let i = 0; i < sanitizedHeaders.length; i++) {
            processedData.headersMap[`header_${i}`] = sanitizedHeaders[i];
        }
        
        // 创建数据行映射
        for (let rowIndex = 0; rowIndex < sanitizedRows.length; rowIndex++) {
            processedData.dataRows[`row_${rowIndex}`] = sanitizedRows[rowIndex];
        }
        
        return processedData;
    }
}

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

// 添加缺失的Excel相关函数
async function loadUserExcelList() {
    showLoading(true);
    try {
        const snapshot = await db.collection('excel_files')
            .orderBy('uploadedAt', 'desc')
            .get();

        const excelList = document.getElementById('userExcelList');
        if (!excelList) return;

        if (snapshot.empty) {
            excelList.innerHTML = `
                <div class="empty-state">
                    <i class="ri-file-excel-line" style="font-size: 48px; color: #666;"></i>
                    <p>暂无Excel文件</p>
                </div>
            `;
            return;
        }

        excelList.innerHTML = snapshot.docs.map(doc => {
            const file = doc.data();
            return `
                <div class="excel-file-card">
                    <div class="file-header">
                        <i class="ri-file-excel-line"></i>
                        <h3>${file.fileName}</h3>
                    </div>
                    <div class="file-info">
                        <p>上传者: ${file.uploadedBy}</p>
                        <p>上传时间: ${formatDate(file.uploadedAt)}</p>
                        <p>数据行数: ${file.totalRows}</p>
                    </div>
                    <div class="file-actions">
                        <button onclick="viewExcelFile('${doc.id}')" class="btn btn-primary">查看</button>
                        <button onclick="downloadExcelData('${doc.id}')" class="btn btn-secondary">下载</button>
                    </div>
                </div>
            `;
        }).join('');

    } catch (error) {
        console.error('加载Excel文件列表失败:', error);
        alert('加载失败: ' + error.message);
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

        const excelList = document.getElementById('adminExcelList');
        if (!excelList) return;

        if (snapshot.empty) {
            excelList.innerHTML = `
                <div class="empty-state">
                    <i class="ri-file-excel-line" style="font-size: 48px; color: #666;"></i>
                    <p>暂无Excel文件</p>
                </div>
            `;
            return;
        }

        excelList.innerHTML = snapshot.docs.map(doc => {
            const file = doc.data();
            return `
                <div class="excel-file-card">
                    <div class="file-header">
                        <i class="ri-file-excel-line"></i>
                        <h3>${file.fileName}</h3>
                    </div>
                    <div class="file-info">
                        <p>上传者: ${file.uploadedBy}</p>
                        <p>上传时间: ${formatDate(file.uploadedAt)}</p>
                        <p>数据行数: ${file.totalRows}</p>
                    </div>
                    <div class="file-actions">
                        <button onclick="viewExcelFile('${doc.id}')" class="btn btn-primary">查看</button>
                        <button onclick="downloadExcelData('${doc.id}')" class="btn btn-secondary">下载</button>
                        <button onclick="deleteExcelFile('${doc.id}')" class="btn btn-danger">删除</button>
                    </div>
                </div>
            `;
        }).join('');

    } catch (error) {
        console.error('加载Excel文件列表失败:', error);
        alert('加载失败: ' + error.message);
    } finally {
        showLoading(false);
    }
}

async function deleteExcelFile(fileId) {
    if (!confirm('确定要删除这个Excel文件吗？此操作不可恢复。')) {
        return;
    }

    showLoading(true);
    try {
        await db.collection('excel_files').doc(fileId).delete();
        alert('Excel文件已删除');
        if (currentUser.role === 'admin') {
            await loadAdminExcelList();
        } else {
            await loadUserExcelList();
        }
    } catch (error) {
        console.error('删除Excel文件失败:', error);
        alert('操作失败: ' + error.message);
    } finally {
        showLoading(false);
    }
}

function clearExcelPreview() {
    const previewDiv = document.getElementById('excelPreview');
    if (previewDiv) {
        previewDiv.style.display = 'none';
    }
    currentExcelData = null;
}

function clearAdminExcelPreview() {
    const previewDiv = document.getElementById('adminExcelPreview');
    if (previewDiv) {
        previewDiv.style.display = 'none';
    }
    currentExcelData = null;
}

function handleLogout() {
    currentUser = null;
    showLoginForm();
}

// 改进的Excel文件读取函数 - 支持多个sheet选择
function readExcelFile(file, isAdmin = false) {
    const reader = new FileReader();
    
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            console.log('📋 Excel工作簿信息:', {
                sheetNames: workbook.SheetNames,
                totalSheets: workbook.SheetNames.length
            });
            
            // 如果有多个sheet，显示选择器
            if (workbook.SheetNames.length > 1) {
                showSheetSelector(workbook, file, isAdmin);
            } else {
                // 只有一个sheet，直接处理
                processWorksheet(workbook, workbook.SheetNames[0], file, isAdmin);
            }
            
        } catch (error) {
            console.error('Excel文件读取失败:', error);
            alert('Excel文件读取失败，请检查文件格式');
        }
    };

    reader.readAsArrayBuffer(file);
}

// 显示工作表选择器
function showSheetSelector(workbook, file, isAdmin) {
    const modal = document.createElement('div');
    modal.className = 'modal';
    modal.style.display = 'block';
    modal.innerHTML = `
        <div class="modal-content">
            <div class="modal-header">
                <h3>选择工作表</h3>
                <button onclick="this.closest('.modal').remove()" class="modal-close">&times;</button>
            </div>
            <div class="modal-body">
                <p>检测到多个工作表，请选择要上传的工作表：</p>
                <div class="sheet-selector">
                    ${workbook.SheetNames.map((sheetName, index) => {
                        const worksheet = workbook.Sheets[sheetName];
                        const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
                        const rowCount = range.e.r - range.s.r;
                        const colCount = range.e.c - range.s.c + 1;
                        
                        return `
                            <div class="sheet-option" onclick="selectSheet('${sheetName}', ${index})">
                                <div class="sheet-info">
                                    <h4>${sheetName}</h4>
                                    <p>行数: ${rowCount}, 列数: ${colCount}</p>
                                </div>
                                <i class="ri-arrow-right-line"></i>
                            </div>
                        `;
                    }).join('')}
                </div>
                <div class="modal-actions">
                    <button onclick="selectAllSheets()" class="btn btn-primary">
                        <i class="ri-stack-line"></i> 全部上传
                    </button>
                    <button onclick="this.closest('.modal').remove()" class="btn btn-secondary">取消</button>
                </div>
            </div>
        </div>
    `;

    document.body.appendChild(modal);
    
    // 保存工作簿信息到全局变量
    window.currentWorkbook = workbook;
    window.currentFile = file;
    window.currentIsAdmin = isAdmin;
}

// 选择单个工作表
function selectSheet(sheetName, index) {
    const workbook = window.currentWorkbook;
    const file = window.currentFile;
    const isAdmin = window.currentIsAdmin;
    
    processWorksheet(workbook, sheetName, file, isAdmin);
    
    // 关闭选择器
    document.querySelector('.modal').remove();
}

// 选择所有工作表
function selectAllSheets() {
    const workbook = window.currentWorkbook;
    const file = window.currentFile;
    const isAdmin = window.currentIsAdmin;
    
    // 处理所有工作表
    const allSheetsData = [];
    
    workbook.SheetNames.forEach(sheetName => {
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
            header: 1,
            defval: '', 
            raw: false  
        });
        
        if (jsonData.length > 0) {
            const cleanHeaders = jsonData[0] ? jsonData[0].map((header, index) => {
                if (header === null || header === undefined || header === '') {
                    return `列${index + 1}`;
                }
                return String(header);
            }) : [];

            const cleanRows = jsonData.slice(1).map(row => {
                if (!Array.isArray(row)) {
                    return new Array(cleanHeaders.length).fill('');
                }
                return row.map(cell => {
                    if (cell === null || cell === undefined) {
                        return '';
                    }
                    if (typeof cell === 'object') {
                        return JSON.stringify(cell);
                    }
                    return String(cell);
                });
            });

            allSheetsData.push({
                sheetName: sheetName,
                headers: cleanHeaders,
                rows: cleanRows
            });
        }
    });

    currentExcelData = {
        fileName: file.name,
        originalFile: file, // 保存原始文件
        fileSize: file.size,
        fileType: file.type,
        lastModified: file.lastModified,
        totalSheets: workbook.SheetNames.length,
        sheetNames: workbook.SheetNames,
        sheets: allSheetsData,
        isMultiSheet: true,
        isAdmin: isAdmin
    };

    console.log('✅ 多工作表Excel数据:', {
        fileName: currentExcelData.fileName,
        totalSheets: currentExcelData.totalSheets,
        sheetNames: currentExcelData.sheetNames,
        sheetsData: currentExcelData.sheets.map(sheet => ({
            name: sheet.sheetName,
            rows: sheet.rows.length,
            cols: sheet.headers.length
        }))
    });

    displayExcelPreview(isAdmin);
    
    // 关闭选择器
    document.querySelector('.modal').remove();
}

// 处理单个工作表
function processWorksheet(workbook, sheetName, file, isAdmin) {
    const worksheet = workbook.Sheets[sheetName];
    
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
        header: 1,
        defval: '', 
        raw: false  
    });
    
    if (jsonData.length === 0) {
        alert('选择的工作表为空');
        return;
    }

    const cleanHeaders = jsonData[0] ? jsonData[0].map((header, index) => {
        if (header === null || header === undefined || header === '') {
            return `列${index + 1}`;
        }
        return String(header);
    }) : [];

    const cleanRows = jsonData.slice(1).map(row => {
        if (!Array.isArray(row)) {
            return new Array(cleanHeaders.length).fill('');
        }
        return row.map(cell => {
            if (cell === null || cell === undefined) {
                return '';
            }
            if (typeof cell === 'object') {
                return JSON.stringify(cell);
            }
            return String(cell);
        });
    });

    currentExcelData = {
        fileName: file.name,
        originalFile: file, // 保存原始文件
        fileSize: file.size,
        fileType: file.type,
        lastModified: file.lastModified,
        sheetName: sheetName,
        totalSheets: workbook.SheetNames.length,
        sheetNames: workbook.SheetNames,
        data: [cleanHeaders, ...cleanRows],
        headers: cleanHeaders,
        rows: cleanRows,
        isMultiSheet: false,
        isAdmin: isAdmin
    };

    console.log('✅ 单工作表Excel数据:', {
        fileName: currentExcelData.fileName,
        sheetName: currentExcelData.sheetName,
        totalSheets: currentExcelData.totalSheets,
        headersCount: currentExcelData.headers.length,
        rowsCount: currentExcelData.rows.length
    });

    displayExcelPreview(isAdmin);
}

// 更新预览显示函数
function displayExcelPreview(isAdmin = false) {
    const previewId = isAdmin ? 'adminExcelPreview' : 'excelPreview';
    const containerId = isAdmin ? 'adminExcelTableContainer' : 'excelTableContainer';
    
    const previewDiv = document.getElementById(previewId);
    const tableContainer = document.getElementById(containerId);
    
    if (!previewDiv || !tableContainer || !currentExcelData) return;

    let tableHTML = `
        <div class="file-info">
            <p><strong>文件名:</strong> ${currentExcelData.fileName}</p>
            <p><strong>文件大小:</strong> ${(currentExcelData.fileSize / 1024).toFixed(2)} KB</p>
            <p><strong>总工作表数:</strong> ${currentExcelData.totalSheets}</p>
    `;

    if (currentExcelData.isMultiSheet) {
        // 多工作表预览
        tableHTML += `<p><strong>将要上传:</strong> 所有工作表 (${currentExcelData.sheets.length}个)</p></div>`;
        
        currentExcelData.sheets.forEach((sheet, index) => {
            if (index < 3) { // 只显示前3个工作表的预览
                tableHTML += `
                    <div class="sheet-preview">
                        <h4>工作表: ${sheet.sheetName}</h4>
                        <p>数据行数: ${sheet.rows.length}, 列数: ${sheet.headers.length}</p>
                        <table class="excel-table">
                            <thead>
                                <tr>
                                    ${sheet.headers.map(header => 
                                        `<th>${String(header || '未命名列')}</th>`
                                    ).join('')}
                                </tr>
                            </thead>
                            <tbody>
                                ${sheet.rows.slice(0, 5).map(row => `
                                    <tr>
                                        ${sheet.headers.map((_, colIndex) => {
                                            const cellValue = row[colIndex];
                                            let displayValue = String(cellValue || '');
                                            return `<td>${displayValue}</td>`;
                                        }).join('')}
                                    </tr>
                                `).join('')}
                                ${sheet.rows.length > 5 ? 
                                    `<tr><td colspan="${sheet.headers.length}" style="text-align: center; color: #666;">... 还有 ${sheet.rows.length - 5} 行</td></tr>` : ''}
                            </tbody>
                        </table>
                    </div>
                `;
            }
        });
        
        if (currentExcelData.sheets.length > 3) {
            tableHTML += `<p style="text-align: center; color: #666;">... 还有 ${currentExcelData.sheets.length - 3} 个工作表</p>`;
        }
    } else {
        // 单工作表预览
        tableHTML += `
            <p><strong>工作表:</strong> ${currentExcelData.sheetName}</p>
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
                ${currentExcelData.rows.slice(0, 10).map(row => `
                    <tr>
                        ${currentExcelData.headers.map((_, index) => {
                            const cellValue = row[index];
                            let displayValue = String(cellValue || '');
                            return `<td>${displayValue}</td>`;
                        }).join('')}
                    </tr>
                `).join('')}
                ${currentExcelData.rows.length > 10 ? 
                    `<tr><td colspan="${currentExcelData.headers.length}" style="text-align: center; color: #666;">... 还有 ${currentExcelData.rows.length - 10} 行数据</td></tr>` : ''}
            </tbody>
        </table>
        `;
    }
    
    tableContainer.innerHTML = tableHTML;
}

// 更新数据处理函数以支持多工作表
function processExcelDataForFirestore(excelData) {
    console.log('开始处理Excel数据，原始数据结构:', excelData);
    
    if (excelData.isMultiSheet) {
        // 处理多工作表数据
        const processedData = {
            fileName: String(excelData.fileName || '未知文件'),
            fileSize: excelData.fileSize,
            fileType: excelData.fileType,
            lastModified: excelData.lastModified,
            totalSheets: excelData.totalSheets,
            sheetNames: excelData.sheetNames.join('|||'),
            isMultiSheet: true,
            sheets: {},
            totalRows: 0
        };
        
        // 处理每个工作表
        excelData.sheets.forEach((sheet, sheetIndex) => {
            const sheetKey = `sheet_${sheetIndex}`;
            const sanitizedHeaders = [];
            
            // 处理headers
            for (let i = 0; i < sheet.headers.length; i++) {
                const header = sheet.headers[i];
                let cleanHeader = String(header || `列${i + 1}`);
                sanitizedHeaders.push(cleanHeader);
            }
            
            // 处理行数据
            const dataRows = {};
            sheet.rows.forEach((row, rowIndex) => {
                const sanitizedRow = {};
                for (let colIndex = 0; colIndex < sanitizedHeaders.length; colIndex++) {
                    const headerKey = `col_${colIndex}`;
                    const cell = row[colIndex];
                    const finalValue = totallyFlattenData(cell);
                    sanitizedRow[headerKey] = finalValue;
                }
                dataRows[`row_${rowIndex}`] = sanitizedRow;
            });
            
            processedData.sheets[sheetKey] = {
                sheetName: sheet.sheetName,
                headersList: sanitizedHeaders.join('|||'),
                headersMap: {},
                dataRows: dataRows,
                totalRows: sheet.rows.length
            };
            
            // 创建headers映射
            for (let i = 0; i < sanitizedHeaders.length; i++) {
                processedData.sheets[sheetKey].headersMap[`header_${i}`] = sanitizedHeaders[i];
            }
            
            processedData.totalRows += sheet.rows.length;
        });
        
        return processedData;
    } else {
        // 处理单工作表数据（原有逻辑）
        const sanitizedHeaders = [];
        if (Array.isArray(excelData.headers)) {
            for (let i = 0; i < excelData.headers.length; i++) {
                const header = excelData.headers[i];
                let cleanHeader = String(header || `列${i + 1}`);
                sanitizedHeaders.push(cleanHeader);
            }
        }
        
        const sanitizedRows = [];
        if (Array.isArray(excelData.rows)) {
            for (let rowIndex = 0; rowIndex < excelData.rows.length; rowIndex++) {
                const row = excelData.rows[rowIndex];
                const sanitizedRow = {};
                
                if (Array.isArray(row)) {
                    for (let colIndex = 0; colIndex < sanitizedHeaders.length; colIndex++) {
                        const headerKey = `col_${colIndex}`;
                        const cell = row[colIndex];
                        const finalValue = totallyFlattenData(cell);
                        sanitizedRow[headerKey] = finalValue;
                    }
                }
                
                sanitizedRows.push(sanitizedRow);
            }
        }
        
        const processedData = {
            fileName: String(excelData.fileName || '未知文件'),
            fileSize: excelData.fileSize,
            fileType: excelData.fileType,
            lastModified: excelData.lastModified,
            sheetName: excelData.sheetName,
            totalSheets: excelData.totalSheets,
            sheetNames: excelData.sheetNames.join('|||'),
            isMultiSheet: false,
            headersList: sanitizedHeaders.join('|||'),
            headersMap: {},
            dataRows: {},
            totalRows: sanitizedRows.length
        };
        
        // 创建headers映射
        for (let i = 0; i < sanitizedHeaders.length; i++) {
            processedData.headersMap[`header_${i}`] = sanitizedHeaders[i];
        }
        
        // 创建数据行映射
        for (let rowIndex = 0; rowIndex < sanitizedRows.length; rowIndex++) {
            processedData.dataRows[`row_${rowIndex}`] = sanitizedRows[rowIndex];
        }
        
        return processedData;
    }
}

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
window.selectSheet = selectSheet;
window.selectAllSheets = selectAllSheets;

// 改进的Excel文件读取函数 - 支持多个sheet选择
function readExcelFile(file, isAdmin = false) {
    const reader = new FileReader();
    
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            console.log('📋 Excel工作簿信息:', {
                sheetNames: workbook.SheetNames,
                totalSheets: workbook.SheetNames.length
            });
            
            // 如果有多个sheet，显示选择器
            if (workbook.SheetNames.length > 1) {
                showSheetSelector(workbook, file, isAdmin);
            } else {
                // 只有一个sheet，直接处理
                processWorksheet(workbook, workbook.SheetNames[0], file, isAdmin);
            }
            
        } catch (error) {
            console.error('Excel文件读取失败:', error);
            alert('Excel文件读取失败，请检查文件格式');
        }
    };

    reader.readAsArrayBuffer(file);
}

// 显示工作表选择器
function showSheetSelector(workbook, file, isAdmin) {
    const modal = document.createElement('div');
    modal.className = 'modal';
    modal.style.display = 'block';
    modal.innerHTML = `
        <div class="modal-content">
            <div class="modal-header">
                <h3>选择工作表</h3>
                <button onclick="this.closest('.modal').remove()" class="modal-close">&times;</button>
            </div>
            <div class="modal-body">
                <p>检测到多个工作表，请选择要上传的工作表：</p>
                <div class="sheet-selector">
                    ${workbook.SheetNames.map((sheetName, index) => {
                        const worksheet = workbook.Sheets[sheetName];
                        const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
                        const rowCount = range.e.r - range.s.r;
                        const colCount = range.e.c - range.s.c + 1;
                        
                        return `
                            <div class="sheet-option" onclick="selectSheet('${sheetName}', ${index})">
                                <div class="sheet-info">
                                    <h4>${sheetName}</h4>
                                    <p>行数: ${rowCount}, 列数: ${colCount}</p>
                                </div>
                                <i class="ri-arrow-right-line"></i>
                            </div>
                        `;
                    }).join('')}
                </div>
                <div class="modal-actions">
                    <button onclick="selectAllSheets()" class="btn btn-primary">
                        <i class="ri-stack-line"></i> 全部上传
                    </button>
                    <button onclick="this.closest('.modal').remove()" class="btn btn-secondary">取消</button>
                </div>
            </div>
        </div>
    `;

    document.body.appendChild(modal);
    
    // 保存工作簿信息到全局变量
    window.currentWorkbook = workbook;
    window.currentFile = file;
    window.currentIsAdmin = isAdmin;
}

// 选择单个工作表
function selectSheet(sheetName, index) {
    const workbook = window.currentWorkbook;
    const file = window.currentFile;
    const isAdmin = window.currentIsAdmin;
    
    processWorksheet(workbook, sheetName, file, isAdmin);
    
    // 关闭选择器
    document.querySelector('.modal').remove();
}

// 选择所有工作表
function selectAllSheets() {
    const workbook = window.currentWorkbook;
    const file = window.currentFile;
    const isAdmin = window.currentIsAdmin;
    
    // 处理所有工作表
    const allSheetsData = [];
    
    workbook.SheetNames.forEach(sheetName => {
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
            header: 1,
            defval: '', 
            raw: false  
        });
        
        if (jsonData.length > 0) {
            const cleanHeaders = jsonData[0] ? jsonData[0].map((header, index) => {
                if (header === null || header === undefined || header === '') {
                    return `列${index + 1}`;
                }
                return String(header);
            }) : [];

            const cleanRows = jsonData.slice(1).map(row => {
                if (!Array.isArray(row)) {
                    return new Array(cleanHeaders.length).fill('');
                }
                return row.map(cell => {
                    if (cell === null || cell === undefined) {
                        return '';
                    }
                    if (typeof cell === 'object') {
                        return JSON.stringify(cell);
                    }
                    return String(cell);
                });
            });

            allSheetsData.push({
                sheetName: sheetName,
                headers: cleanHeaders,
                rows: cleanRows
            });
        }
    });

    currentExcelData = {
        fileName: file.name,
        originalFile: file, // 保存原始文件
        fileSize: file.size,
        fileType: file.type,
        lastModified: file.lastModified,
        totalSheets: workbook.SheetNames.length,
        sheetNames: workbook.SheetNames,
        sheets: allSheetsData,
        isMultiSheet: true,
        isAdmin: isAdmin
    };

    console.log('✅ 多工作表Excel数据:', {
        fileName: currentExcelData.fileName,
        totalSheets: currentExcelData.totalSheets,
        sheetNames: currentExcelData.sheetNames,
        sheetsData: currentExcelData.sheets.map(sheet => ({
            name: sheet.sheetName,
            rows: sheet.rows.length,
            cols: sheet.headers.length
        }))
    });

    displayExcelPreview(isAdmin);
    
    // 关闭选择器
    document.querySelector('.modal').remove();
}

// 处理单个工作表
function processWorksheet(workbook, sheetName, file, isAdmin) {
    const worksheet = workbook.Sheets[sheetName];
    
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
        header: 1,
        defval: '', 
        raw: false  
    });
    
    if (jsonData.length === 0) {
        alert('选择的工作表为空');
        return;
    }

    const cleanHeaders = jsonData[0] ? jsonData[0].map((header, index) => {
        if (header === null || header === undefined || header === '') {
            return `列${index + 1}`;
        }
        return String(header);
    }) : [];

    const cleanRows = jsonData.slice(1).map(row => {
        if (!Array.isArray(row)) {
            return new Array(cleanHeaders.length).fill('');
        }
        return row.map(cell => {
            if (cell === null || cell === undefined) {
                return '';
            }
            if (typeof cell === 'object') {
                return JSON.stringify(cell);
            }
            return String(cell);
        });
    });

    currentExcelData = {
        fileName: file.name,
        originalFile: file, // 保存原始文件
        fileSize: file.size,
        fileType: file.type,
        lastModified: file.lastModified,
        sheetName: sheetName,
        totalSheets: workbook.SheetNames.length,
        sheetNames: workbook.SheetNames,
        data: [cleanHeaders, ...cleanRows],
        headers: cleanHeaders,
        rows: cleanRows,
        isMultiSheet: false,
        isAdmin: isAdmin
    };

    console.log('✅ 单工作表Excel数据:', {
        fileName: currentExcelData.fileName,
        sheetName: currentExcelData.sheetName,
        totalSheets: currentExcelData.totalSheets,
        headersCount: currentExcelData.headers.length,
        rowsCount: currentExcelData.rows.length
    });

    displayExcelPreview(isAdmin);
}

// 更新预览显示函数
function displayExcelPreview(isAdmin = false) {
    const previewId = isAdmin ? 'adminExcelPreview' : 'excelPreview';
    const containerId = isAdmin ? 'adminExcelTableContainer' : 'excelTableContainer';
    
    const previewDiv = document.getElementById(previewId);
    const tableContainer = document.getElementById(containerId);
    
    if (!previewDiv || !tableContainer || !currentExcelData) return;

    let tableHTML = `
        <div class="file-info">
            <p><strong>文件名:</strong> ${currentExcelData.fileName}</p>
            <p><strong>文件大小:</strong> ${(currentExcelData.fileSize / 1024).toFixed(2)} KB</p>
            <p><strong>总工作表数:</strong> ${currentExcelData.totalSheets}</p>
    `;

    if (currentExcelData.isMultiSheet) {
        // 多工作表预览
        tableHTML += `<p><strong>将要上传:</strong> 所有工作表 (${currentExcelData.sheets.length}个)</p></div>`;
        
        currentExcelData.sheets.forEach((sheet, index) => {
            if (index < 3) { // 只显示前3个工作表的预览
                tableHTML += `
                    <div class="sheet-preview">
                        <h4>工作表: ${sheet.sheetName}</h4>
                        <p>数据行数: ${sheet.rows.length}, 列数: ${sheet.headers.length}</p>
                        <table class="excel-table">
                            <thead>
                                <tr>
                                    ${sheet.headers.map(header => 
                                        `<th>${String(header || '未命名列')}</th>`
                                    ).join('')}
                                </tr>
                            </thead>
                            <tbody>
                                ${sheet.rows.slice(0, 5).map(row => `
                                    <tr>
                                        ${sheet.headers.map((_, colIndex) => {
                                            const cellValue = row[colIndex];
                                            let displayValue = String(cellValue || '');
                                            return `<td>${displayValue}</td>`;
                                        }).join('')}
                                    </tr>
                                `).join('')}
                                ${sheet.rows.length > 5 ? 
                                    `<tr><td colspan="${sheet.headers.length}" style="text-align: center; color: #666;">... 还有 ${sheet.rows.length - 5} 行</td></tr>` : ''}
                            </tbody>
                        </table>
                    </div>
                `;
            }
        });
        
        if (currentExcelData.sheets.length > 3) {
            tableHTML += `<p style="text-align: center; color: #666;">... 还有 ${currentExcelData.sheets.length - 3} 个工作表</p>`;
        }
    } else {
        // 单工作表预览
        tableHTML += `
            <p><strong>工作表:</strong> ${currentExcelData.sheetName}</p>
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
                ${currentExcelData.rows.slice(0, 10).map(row => `
                    <tr>
                        ${currentExcelData.headers.map((_, index) => {
                            const cellValue = row[index];
                            let displayValue = String(cellValue || '');
                            return `<td>${displayValue}</td>`;
                        }).join('')}
                    </tr>
                `).join('')}
                ${currentExcelData.rows.length > 10 ? 
                    `<tr><td colspan="${currentExcelData.headers.length}" style="text-align: center; color: #666;">... 还有 ${currentExcelData.rows.length - 10} 行数据</td></tr>` : ''}
            </tbody>
        </table>
        `;
    }
    
    tableContainer.innerHTML = tableHTML;
}

// 更新数据处理函数以支持多工作表
function processExcelDataForFirestore(excelData) {
    console.log('开始处理Excel数据，原始数据结构:', excelData);
    
    if (excelData.isMultiSheet) {
        // 处理多工作表数据
        const processedData = {
            fileName: String(excelData.fileName || '未知文件'),
            fileSize: excelData.fileSize,
            fileType: excelData.fileType,
            lastModified: excelData.lastModified,
            totalSheets: excelData.totalSheets,
            sheetNames: excelData.sheetNames.join('|||'),
            isMultiSheet: true,
            sheets: {},
            totalRows: 0
        };
        
        // 处理每个工作表
        excelData.sheets.forEach((sheet, sheetIndex) => {
            const sheetKey = `sheet_${sheetIndex}`;
            const sanitizedHeaders = [];
            
            // 处理headers
            for (let i = 0; i < sheet.headers.length; i++) {
                const header = sheet.headers[i];
                let cleanHeader = String(header || `列${i + 1}`);
                sanitizedHeaders.push(cleanHeader);
            }
            
            // 处理行数据
            const dataRows = {};
            sheet.rows.forEach((row, rowIndex) => {
                const sanitizedRow = {};
                for (let colIndex = 0; colIndex < sanitizedHeaders.length; colIndex++) {
                    const headerKey = `col_${colIndex}`;
                    const cell = row[colIndex];
                    const finalValue = totallyFlattenData(cell);
                    sanitizedRow[headerKey] = finalValue;
                }
                dataRows[`row_${rowIndex}`] = sanitizedRow;
            });
            
            processedData.sheets[sheetKey] = {
                sheetName: sheet.sheetName,
                headersList: sanitizedHeaders.join('|||'),
                headersMap: {},
                dataRows: dataRows,
                totalRows: sheet.rows.length
            };
            
            // 创建headers映射
            for (let i = 0; i < sanitizedHeaders.length; i++) {
                processedData.sheets[sheetKey].headersMap[`header_${i}`] = sanitizedHeaders[i];
            }
            
            processedData.totalRows += sheet.rows.length;
        });
        
        return processedData;
    } else {
        // 处理单工作表数据（原有逻辑）
        const sanitizedHeaders = [];
        if (Array.isArray(excelData.headers)) {
            for (let i = 0; i < excelData.headers.length; i++) {
                const header = excelData.headers[i];
                let cleanHeader = String(header || `列${i + 1}`);
                sanitizedHeaders.push(cleanHeader);
            }
        }
        
        const sanitizedRows = [];
        if (Array.isArray(excelData.rows)) {
            for (let rowIndex = 0; rowIndex < excelData.rows.length; rowIndex++) {
                const row = excelData.rows[rowIndex];
                const sanitizedRow = {};
                
                if (Array.isArray(row)) {
                    for (let colIndex = 0; colIndex < sanitizedHeaders.length; colIndex++) {
                        const headerKey = `col_${colIndex}`;
                        const cell = row[colIndex];
                        const finalValue = totallyFlattenData(cell);
                        sanitizedRow[headerKey] = finalValue;
                    }
                }
                
                sanitizedRows.push(sanitizedRow);
            }
        }
        
        const processedData = {
            fileName: String(excelData.fileName || '未知文件'),
            fileSize: excelData.fileSize,
            fileType: excelData.fileType,
            lastModified: excelData.lastModified,
            sheetName: excelData.sheetName,
            totalSheets: excelData.totalSheets,
            sheetNames: excelData.sheetNames.join('|||'),
            isMultiSheet: false,
            headersList: sanitizedHeaders.join('|||'),
            headersMap: {},
            dataRows: {},
            totalRows: sanitizedRows.length
        };
        
        // 创建headers映射
        for (let i = 0; i < sanitizedHeaders.length; i++) {
            processedData.headersMap[`header_${i}`] = sanitizedHeaders[i];
        }
        
        // 创建数据行映射
        for (let rowIndex = 0; rowIndex < sanitizedRows.length; rowIndex++) {
            processedData.dataRows[`row_${rowIndex}`] = sanitizedRows[rowIndex];
        }
        
        return processedData;
    }
}
