<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Access 数据库管理工具</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Microsoft YaHei', Arial, sans-serif; background: #f5f5f5; padding: 20px; }
        .container { max-width: 1400px; margin: 0 auto; background: white; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); overflow: hidden; }
        .header { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 20px 30px; }
        .header h1 { font-size: 24px; margin-bottom: 10px; }
        .header p { opacity: 0.9; font-size: 14px; }
        .tabs { display: flex; background: #f8f9fa; border-bottom: 2px solid #dee2e6; }
        .tab { padding: 15px 25px; cursor: pointer; border: none; background: none; font-size: 14px; color: #495057; border-bottom: 3px solid transparent; transition: all 0.3s; }
        .tab:hover { background: #e9ecef; }
        .tab.active { color: #667eea; border-bottom-color: #667eea; background: white; font-weight: bold; }
        .content { padding: 30px; }
        .panel { display: none; }
        .panel.active { display: block; }
        .form-group { margin-bottom: 20px; }
        .form-group label { display: block; margin-bottom: 8px; font-weight: 500; color: #333; }
        .form-group input, .form-group textarea { width: 100%; padding: 10px; border: 1px solid #ddd; border-radius: 4px; font-size: 14px; }
        .form-group textarea { height: 100px; font-family: 'Consolas', monospace; resize: vertical; }
        .btn { padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; font-size: 14px; transition: all 0.3s; }
        .btn-primary { background: #667eea; color: white; }
        .btn-primary:hover { background: #5568d3; }
        .btn-success { background: #28a745; color: white; }
        .btn-success:hover { background: #218838; }
        .btn-danger { background: #dc3545; color: white; }
        .btn-danger:hover { background: #c82333; }
        .btn-info { background: #17a2b8; color: white; }
        .btn-info:hover { background: #138496; }
        table { width: 100%; border-collapse: collapse; margin-top: 15px; }
        th, td { padding: 12px; text-align: left; border-bottom: 1px solid #dee2e6; }
        th { background: #f8f9fa; font-weight: 600; color: #495057; }
        tr:hover { background: #f8f9fa; }
        .alert { padding: 15px; border-radius: 4px; margin-bottom: 20px; }
        .alert-success { background: #d4edda; color: #155724; border: 1px solid #c3e6cb; }
        .alert-danger { background: #f8d7da; color: #721c24; border: 1px solid #f5c6cb; }
        .alert-info { background: #d1ecf1; color: #0c5460; border: 1px solid #bee5eb; }
        .table-list { display: grid; grid-template-columns: repeat(auto-fill, minmax(200px, 1fr)); gap: 15px; }
        .table-card { background: white; border: 2px solid #dee2e6; border-radius: 8px; padding: 15px; cursor: pointer; transition: all 0.3s; }
        .table-card:hover { border-color: #667eea; box-shadow: 0 4px 12px rgba(102, 126, 234, 0.2); }
        .table-card h3 { font-size: 16px; margin-bottom: 5px; color: #333; }
        .table-card p { font-size: 12px; color: #6c757d; }
        .action-bar { display: flex; gap: 10px; margin-bottom: 20px; flex-wrap: wrap; }
        .status-bar { display: flex; justify-content: space-between; align-items: center; padding: 10px 15px; background: #f8f9fa; border-radius: 4px; margin-bottom: 15px; font-size: 13px; }
        .pagination { display: flex; gap: 5px; margin-top: 15px; }
        .pagination button { padding: 5px 10px; border: 1px solid #dee2e6; background: white; cursor: pointer; }
        .pagination button:hover { background: #e9ecef; }
        .pagination button.active { background: #667eea; color: white; border-color: #667eea; }
        .code-block { background: #282c34; color: #abb2bf; padding: 15px; border-radius: 4px; overflow-x: auto; font-family: 'Consolas', monospace; font-size: 13px; }
        #dataTable tbody tr:hover { background: #e3f2fd; }
        #dataTable tbody tr { transition: background 0.2s; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header" style="display: flex; justify-content: space-between; align-items: center;">
            <div>
                <h1>📊 Access 数据库管理工具</h1>
                <p>简单易用的数据库管理界面</p>
            </div>
            <a href="?action=logout" style="background: rgba(255,255,255,0.2); color: white; padding: 10px 20px; text-decoration: none; border-radius: 4px; font-size: 14px;">退出登录</a>
        </div>
        
        <div class="tabs">
            <button class="tab active" onclick="showPanel('connect')">🔗 连接数据库</button>
            <button class="tab" onclick="showPanel('tables')">📋 表列表</button>
            <button class="tab" onclick="showPanel('query')">🔍 SQL查询</button>
            <button class="tab" onclick="showPanel('structure')">🏗️ 表结构</button>
        </div>
        
        <div class="content">
            <%
            ' 获取当前面板
            Dim currentPanel
            currentPanel = Request.QueryString("panel")
            If currentPanel = "" Then currentPanel = "connect"
            
            ' 数据库连接字符串
            Dim dbPath, connectionString, conn
            dbPath = Request.Form("dbpath")
            If dbPath = "" Then dbPath = Request.Cookies("dbpath")
            If dbPath = "" Then dbPath = "database.mdb"
            
            ' 检查是否为相对路径（不包含驱动器号）
            If InStr(dbPath, ":") = 0 Then
                dbPath = Server.MapPath(dbPath)
            End If
            
            connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbPath & ";"
            
            ' ========== 密码认证 ==========
            ' 管理密码（请修改为您自己的密码）
            Dim adminPassword
            adminPassword = "admin888"  ' 在这里修改您的密码
            
            Dim isAuthenticated, authMessage
            isAuthenticated = False
            authMessage = ""
            
            ' 检查是否已登录
            If Request.Cookies("auth_token") = adminPassword Then
                isAuthenticated = True
            End If
            
            ' 处理登录请求
            If Request.ServerVariables("REQUEST_METHOD") = "POST" And Request.Form("action") = "login" Then
                Dim inputPassword
                inputPassword = Request.Form("password")
                
                If inputPassword = adminPassword Then
                    isAuthenticated = True
                    Response.Cookies("auth_token") = adminPassword
                    Response.Cookies("auth_token").Expires = Now() + 1  ' 1天后过期
                Else
                    authMessage = "密码错误，请重试！"
                End If
            End If
            
            ' 处理登出请求
            If Request.QueryString("action") = "logout" Then
                Response.Cookies("auth_token") = ""
                Response.Cookies("auth_token").Expires = Now() - 365  ' 立即过期
                isAuthenticated = False
            End If
            ' ========== 认证结束 ==========
            
            ' 数据库连接相关变量初始化
            Dim isConnected, connError
            Dim sqlQuery, queryResult, queryError
            isConnected = False
            connError = ""
            sqlQuery = ""
            queryResult = ""
            queryError = ""
            
            ' 只有在认证通过后才连接数据库
            If isAuthenticated Then
                ' 检查是否连接成功
                On Error Resume Next
                Set conn = Server.CreateObject("ADODB.Connection")
                conn.Open connectionString
                If Err.Number = 0 Then
                    isConnected = True
                    Response.Cookies("dbpath") = dbPath
                    Response.Cookies("dbpath").Expires = DateAdd("d", 30)
                Else
                    connError = Err.Description
                End If
                On Error GoTo 0
                
                ' 处理SQL查询
                sqlQuery = Request.Form("sqlquery")
                
                ' 检查是否是 Base64 编码（简单判断：包含字母数字和+/=，不包含空格和换行）
                If InStr(sqlQuery, " ") = 0 And InStr(sqlQuery, vbCrLf) = 0 And Len(sqlQuery) > 0 Then
                    ' 尝试 Base64 解码
                    On Error Resume Next
                    Dim decodedQuery
                    decodedQuery = Base64Decode(sqlQuery)
                    If Err.Number = 0 Then
                        sqlQuery = decodedQuery
                    End If
                    On Error GoTo 0
                End If
                
                If Request.ServerVariables("REQUEST_METHOD") = "POST" And sqlQuery <> "" And isConnected Then
                On Error Resume Next
                
                ' 检查是否是 SELECT 查询
                Dim sqlType
                sqlType = UCase(Left(Trim(sqlQuery), 6))
                
                Dim recordsAffected
                recordsAffected = 0
                
                If sqlType = "SELECT" Then
                    Dim rs, output
                    Set rs = conn.Execute(sqlQuery, recordsAffected)
                    If Err.Number = 0 Then
                        If Not rs.EOF Then
                            output = "<table><thead><tr>"
                            For i = 0 To rs.Fields.Count - 1
                                output = output & "<th>" & rs.Fields(i).Name & "</th>"
                            Next
                            output = output & "</tr></thead><tbody>"
                            Dim rowCount
                            rowCount = 0
                            Do While Not rs.EOF And rowCount < 1000
                                output = output & "<tr>"
                                For i = 0 To rs.Fields.Count - 1
                                    Dim fieldValue
                                    fieldValue = rs.Fields(i).Value
                                    If IsNull(fieldValue) Then
                                        fieldValue = "<em style='color:#999'>NULL</em>"
                                    Else
                                        fieldValue = Server.HTMLEncode(CStr(fieldValue))
                                    End If
                                    output = output & "<td>" & fieldValue & "</td>"
                                Next
                                output = output & "</tr>"
                                rowCount = rowCount + 1
                                rs.MoveNext
                            Loop
                            output = output & "</tbody></table>"
                            queryResult = output
                        Else
                            queryResult = "<div class='alert alert-info'>查询成功，但未返回数据。</div>"
                        End If
                        rs.Close
                        Set rs = Nothing
                    Else
                        queryError = Err.Description
                    End If
                Else
                    ' 执行 UPDATE, INSERT, DELETE 等操作
                    conn.Execute sqlQuery, recordsAffected
                    If Err.Number = 0 Then
                        queryResult = "<div class='alert alert-success'>✅ 操作成功！受影响行数: " & recordsAffected & "</div>"
                    Else
                        queryError = Err.Description
                    End If
                End If
                On Error GoTo 0
                End If
            End If
            %>
            
            <!-- 认证界面 -->
            <% If Not isAuthenticated Then %>
                <div style="max-width: 400px; margin: 100px auto; text-align: center;">
                    <h2 style="margin-bottom: 30px; color: #333;">🔐 管理员登录</h2>
                    <% If authMessage <> "" Then %>
                        <div class="alert alert-danger" style="margin-bottom: 20px;"><%=authMessage%></div>
                    <% End If %>
                    <form method="post">
                        <div class="form-group">
                            <input type="password" name="password" placeholder="请输入管理员密码" required style="padding: 15px; font-size: 16px;">
                        </div>
                        <input type="hidden" name="action" value="login">
                        <button type="submit" class="btn btn-primary" style="width: 100%; padding: 15px; font-size: 16px;">登录</button>
                    </form>
                    <p style="margin-top: 20px; color: #999; font-size: 12px;">Access 数据库管理工具 | 需要管理员权限</p>
                </div>
            <% Else %>
                <!-- 已认证，显示管理界面 -->
            
            <!-- 连接面板 -->
            <div class="panel <%=IIf(currentPanel="connect", "active", "")%>" id="panel-connect">
                <% If isConnected Then %>
                    <div class="alert alert-success">
                        ✅ 数据库连接成功！<br>
                        数据库路径: <%=dbPath%>
                    </div>
                <% Else %>
                    <% If dbPath <> "" Then %>
                        <div class="alert alert-danger">
                            ❌ 数据库连接失败: <%=connError%>
                        </div>
                    <% End If %>
                <% End If %>
                
                <form method="post">
                    <div class="form-group">
                        <label>数据库路径 (相对路径或绝对路径):</label>
                        <input type="text" name="dbpath" value="<%=dbPath%>" placeholder="例如: database.mdb 或 C:\data\database.mdb">
                    </div>
                    <div class="form-group">
                        <label>示例路径:</label>
                        <ul style="margin-left: 20px; color: #666; font-size: 13px;">
                            <li>相对路径: database.mdb</li>
                            <li>Server.MapPath: <%=Server.MapPath("database.mdb")%></li>
                        </ul>
                    </div>
                    <button type="submit" class="btn btn-primary" name="action" value="connect">连接数据库</button>
                </form>
            </div>
            
            <!-- 表列表面板 -->
            <div class="panel <%=IIf(currentPanel="tables", "active", "")%>" id="panel-tables">
                <% If isConnected Then %>
                    <div class="status-bar">
                        <span>📁 当前数据库: <%=dbPath%></span>
                        <span>🔗 连接状态: 已连接</span>
                    </div>
                    
                    <%
                    Dim schemaRS
                    Set schemaRS = conn.OpenSchema(20) ' adSchemaTables
                    Dim tableCount
                    tableCount = 0
                    %>
                    
                    <div class="table-list">
                        <%
                        Do While Not schemaRS.EOF
                            Dim tableName, tableType
                            tableName = schemaRS("TABLE_NAME")
                            tableType = schemaRS("TABLE_TYPE")
                            
                            If tableType = "TABLE" Then
                                tableCount = tableCount + 1
                        %>
                        <div class="table-card" onclick="viewTable('<%=tableName%>')">
                            <h3>📄 <%=tableName%></h3>
                            <p>点击查看数据</p>
                        </div>
                        <%
                            End If
                            schemaRS.MoveNext
                        Loop
                        schemaRS.Close
                        Set schemaRS = Nothing
                        %>
                    </div>
                    
                    <% If tableCount = 0 Then %>
                        <div class="alert alert-info">数据库中没有找到表</div>
                    <% Else %>
                        <div class="alert alert-info">共找到 <%=tableCount%> 个表</div>
                    <% End If %>
                <% Else %>
                    <div class="alert alert-danger">请先连接数据库</div>
                <% End If %>
            </div>
            
            <!-- SQL查询面板 -->
            <div class="panel <%=IIf(currentPanel="query", "active", "")%>" id="panel-query">
                <% If isConnected Then %>
                    <% If queryError <> "" Then %>
                        <div class="alert alert-danger">
                            ❌ 查询错误: <%=queryError%>
                        </div>
                    <% End If %>
                    
                    <% If queryResult <> "" Then %>
                        <div class="alert alert-success">✅ 查询执行成功</div>
                        <div class="action-bar">
                            <button class="btn btn-info" onclick="showPanel('tables')">返回表列表</button>
                        </div>
                        <%=queryResult%>
                    <% End If %>
                    
                    <form method="post" id="queryForm" onsubmit="return submitQuery()">
                        <div class="form-group">
                            <label>SQL 语句 (明文 - 不会提交):</label>
                            <textarea id="sqlText" placeholder="输入 SQL 查询语句，例如:
SELECT * FROM 表名
或
SELECT COUNT(*) FROM 表名" onchange="updateBase64()"><%=sqlQuery%></textarea>
                        </div>
                        <div class="form-group">
                            <label>Base64 编码 (自动生成，提交时使用):</label>
                            <textarea id="sqlBase64" name="sqlquery" readonly style="background: #f0f0f0; color: #666;"></textarea>
                        </div>
                        <div class="action-bar">
                            <button type="button" class="btn btn-info" onclick="encodeBase64()">🔒 编码为 Base64</button>
                            <button type="button" class="btn btn-info" onclick="decodeBase64()">🔓 解码 Base64</button>
                            <button type="button" class="btn btn-info" onclick="copyBase64()">📋 复制 Base64</button>
                        </div>
                        <div class="form-group">
                            <label>常用 SQL 示例:</label>
                            <div class="code-block">
SELECT * FROM 表名
SELECT TOP 100 * FROM 表名
SELECT COUNT(*) FROM 表名
SELECT * FROM 表名 WHERE 条件
INSERT INTO 表名 (字段1, 字段2) VALUES (值1, 值2)
UPDATE 表名 SET 字段 = 值 WHERE 条件
DELETE FROM 表名 WHERE 条件
                            </div>
                        </div>
                        <button type="submit" class="btn btn-success">执行查询</button>
                    </form>
                <% Else %>
                    <div class="alert alert-danger">请先连接数据库</div>
                <% End If %>
            </div>
            
            <!-- 表结构面板 -->
            <div class="panel <%=IIf(currentPanel="structure", "active", "")%>" id="panel-structure">
                <% If isConnected Then %>
                    <%
                    Dim selectedTable
                    selectedTable = Request.QueryString("table")
                    
                    If selectedTable <> "" Then
                    %>
                        <div class="status-bar">
                            <span>📋 当前表: <%=selectedTable%></span>
                            <button class="btn btn-info" onclick="showPanel('tables')">返回表列表</button>
                        </div>
                        
                        <%
                        ' 获取表结构
                        Dim columnsRS
                        Set columnsRS = conn.OpenSchema(4, Array(Empty, Empty, selectedTable)) ' adSchemaColumns
                        %>
                        
                        <h3>📊 表结构</h3>
                        <table>
                            <thead>
                                <tr>
                                    <th>字段名</th>
                                    <th>数据类型</th>
                                    <th>大小</th>
                                    <th>允许NULL</th>
                                    <th>主键</th>
                                </tr>
                            </thead>
                            <tbody>
                                <%
                                Do While Not columnsRS.EOF
                                    Dim fieldName, dataType, fieldSize, isNullable, isPrimary
                                    fieldName = columnsRS("COLUMN_NAME")
                                    dataType = columnsRS("DATA_TYPE")
                                    fieldSize = columnsRS("CHARACTER_MAXIMUM_LENGTH")
                                    isNullable = columnsRS("IS_NULLABLE")
                                    
                                    ' 数据类型映射
                                    Dim typeName
                                    Select Case dataType
                                        Case 2: typeName = "SmallInt"
                                        Case 3: typeName = "Integer"
                                        Case 4: typeName = "Single"
                                        Case 5: typeName = "Double"
                                        Case 6: typeName = "Currency"
                                        Case 7: typeName = "Date"
                                        Case 11: typeName = "Boolean"
                                        Case 17: typeName = "Byte"
                                        Case 202: typeName = "VarChar"
                                        Case 203: typeName = "VarWChar"
                                        Case 130: typeName = "WChar"
                                        Case 131: typeName = "Numeric"
                                        Case 135: typeName = "DateTime"
                                        Case Else: typeName = "Type " & dataType
                                    End Select
                                    
                                    If IsNull(fieldSize) Or fieldSize = -1 Then
                                        fieldSize = "-"
                                    End If
                                    
                                    isPrimary = "否"
                                %>
                                <tr>
                                    <td><strong><%=fieldName%></strong></td>
                                    <td><%=typeName%></td>
                                    <td><%=fieldSize%></td>
                                    <td><%=IIf(isNullable="YES", "是", "否")%></td>
                                    <td><%=isPrimary%></td>
                                </tr>
                                <%
                                    columnsRS.MoveNext
                                Loop
                                columnsRS.Close
                                Set columnsRS = Nothing
                                %>
                            </tbody>
                        </table>
                        
                        <h3 style="margin-top: 30px;">📄 数据预览 (前 100 条) <span style="font-size: 14px; color: #666; font-weight: normal;">(点击行进行编辑)</span></h3>
                        <%
                        Dim dataRS
                        Set dataRS = conn.Execute("SELECT TOP 100 * FROM [" & selectedTable & "]")
                        
                        If Not dataRS.EOF Then
                            ' 获取主键字段
                            Dim primaryKeyField
                            primaryKeyField = ""
                            Dim pkRS
                            Set pkRS = conn.OpenSchema(28, Array(Empty, Empty, selectedTable)) ' adSchemaPrimaryKeys
                            If Not pkRS.EOF Then
                                primaryKeyField = pkRS("COLUMN_NAME")
                            End If
                            pkRS.Close
                            Set pkRS = Nothing
                            
                            ' 如果没有找到主键，使用第一个字段
                            If primaryKeyField = "" Then
                                primaryKeyField = dataRS.Fields(0).Name
                            End If
                        %>
                        <table id="dataTable">
                            <thead>
                                <tr>
                                    <% For i = 0 To dataRS.Fields.Count - 1 %>
                                    <th><%=dataRS.Fields(i).Name%></th>
                                    <% Next %>
                                </tr>
                            </thead>
                            <tbody>
                                <%
                                Dim rowId
                                rowId = 0
                                Do While Not dataRS.EOF
                                    rowId = rowId + 1
                                %>
                                <tr onclick="editRow('<%=selectedTable%>', '<%=primaryKeyField%>', <%=rowId%>)" style="cursor: pointer;">
                                    <% For i = 0 To dataRS.Fields.Count - 1 %>
                                    <td>
                                        <% 
                                        Dim val
                                        val = dataRS.Fields(i).Value
                                        If IsNull(val) Then
                                            Response.Write "<em style='color:#999'>NULL</em>"
                                        Else
                                            Response.Write Server.HTMLEncode(CStr(val))
                                        End If
                                        %>
                                    </td>
                                    <% Next %>
                                </tr>
                                <%
                                    dataRS.MoveNext
                                Loop
                                %>
                            </tbody>
                        </table>
                        <script>
                        var rowData = {
                        <%
                        dataRS.MoveFirst
                        rowId = 0
                        Do While Not dataRS.EOF
                            rowId = rowId + 1
                            Response.Write """" & rowId & """: {"
                            For i = 0 To dataRS.Fields.Count - 1
                                Dim jsFieldName, jsFieldValue
                                jsFieldName = dataRS.Fields(i).Name
                                jsFieldValue = dataRS.Fields(i).Value
                                If IsNull(jsFieldValue) Then
                                    jsFieldValue = ""
                                Else
                                    jsFieldValue = Replace(CStr(jsFieldValue), """", "\""")
                                End If
                                Response.Write """" & jsFieldName & """: """ & jsFieldValue & """"
                                If i < dataRS.Fields.Count - 1 Then Response.Write ", "
                            Next
                            Response.Write "}"
                            If Not dataRS.EOF Then Response.Write ", "
                            dataRS.MoveNext
                        Loop
                        %>
                        };
                        </script>
                        <%
                        Else
                        %>
                        <div class="alert alert-info">表中没有数据</div>
                        <%
                        End If
                        dataRS.Close
                        Set dataRS = Nothing
                        %>
                        
                        <h3 style="margin-top: 30px;">🔍 快速查询</h3>
                        <form method="post">
                            <div class="form-group">
                                <textarea name="sqlquery" placeholder="SELECT * FROM <%=selectedTable%> WHERE ...">SELECT TOP 100 * FROM [<%=selectedTable%>]</textarea>
                            </div>
                            <button type="submit" class="btn btn-success" onclick="showPanel('query')">执行查询</button>
                        </form>
                        
                    <%
                    Else
                    %>
                        <div class="alert alert-info">请从表列表中选择一个表查看结构</div>
                        <button class="btn btn-info" onclick="showPanel('tables')">查看表列表</button>
                    <%
                    End If
                    %>
                <% Else %>
                    <div class="alert alert-danger">请先连接数据库</div>
                <% End If %>
            </div>
        </div>
        
        <div style="padding: 20px 30px; background: #f8f9fa; border-top: 1px solid #dee2e6; font-size: 12px; color: #6c757d;">
            Access 数据库管理工具 | 支持 mdb 格式数据库 | 纯 ASP 实现
        </div>
    </div>
    
    <!-- 编辑模态框 -->
    <div id="editModal" style="display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.5); z-index: 1000;">
        <div style="position: fixed; top: 50%; left: 50%; transform: translate(-50%, -50%); background: white; border-radius: 8px; padding: 30px; max-width: 600px; width: 90%; max-height: 80vh; overflow-y: auto; box-shadow: 0 4px 20px rgba(0,0,0,0.3);">
            <h3 style="margin-bottom: 20px;">✏️ 编辑数据</h3>
            <form id="editForm">
                <div id="editFields"></div>
                <div style="margin-top: 20px; text-align: right;">
                    <button type="button" class="btn btn-primary" onclick="saveEdit()">保存</button>
                    <button type="button" class="btn btn-danger" onclick="closeModal()">取消</button>
                </div>
            </form>
        </div>
    </div>
    
    <script>
        function showPanel(panelName) {
            // 隐藏所有面板
            var panels = document.querySelectorAll('.panel');
            panels.forEach(function(panel) {
                panel.classList.remove('active');
            });
            
            // 移除所有标签的active状态
            var tabs = document.querySelectorAll('.tab');
            tabs.forEach(function(tab) {
                tab.classList.remove('active');
            });
            
            // 显示选中的面板
            document.getElementById('panel-' + panelName).classList.add('active');
            
            // 激活对应的标签
            event.target.classList.add('active');
            
            // 更新URL
            var url = new URL(window.location);
            url.searchParams.set('panel', panelName);
            window.history.pushState({}, '', url);
        }
        
        function viewTable(tableName) {
            var url = new URL(window.location);
            url.searchParams.set('panel', 'structure');
            url.searchParams.set('table', tableName);
            window.location.href = url.toString();
        }
        
        // 初始化当前面板
        document.addEventListener('DOMContentLoaded', function() {
            var currentPanel = '<%=currentPanel%>';
            var panels = document.querySelectorAll('.panel');
            var tabs = document.querySelectorAll('.tab');
            
            panels.forEach(function(panel) {
                panel.classList.remove('active');
            });
            
            tabs.forEach(function(tab) {
                tab.classList.remove('active');
            });
            
            var targetPanel = document.getElementById('panel-' + currentPanel);
            if (targetPanel) {
                targetPanel.classList.add('active');
            }
            
            tabs.forEach(function(tab) {
                if (tab.textContent.includes(getPanelText(currentPanel))) {
                    tab.classList.add('active');
                }
            });
        });
        
        function getPanelText(panel) {
            switch(panel) {
                case 'connect': return '连接数据库';
                case 'tables': return '表列表';
                case 'query': return 'SQL查询';
                case 'structure': return '表结构';
                default: return '';
            }
        }
        
        var currentEditTable = '';
        var currentEditPk = '';
        var currentRowId = 0;
        
        function editRow(tableName, pkField, rowId) {
            currentEditTable = tableName;
            currentEditPk = pkField;
            currentRowId = rowId;
            
            var data = rowData[rowId];
            var html = '';
            
            for (var field in data) {
                var value = data[field] || '';
                var isPk = (field === pkField);
                var disabled = isPk ? 'disabled' : '';
                var label = isPk ? field + ' (主键)' : field;
                
                html += '<div class="form-group">';
                html += '<label>' + label + ':</label>';
                html += '<input type="text" name="' + field + '" value="' + value.replace(/"/g, '&quot;') + '" ' + disabled + '>';
                html += '</div>';
            }
            
            document.getElementById('editFields').innerHTML = html;
            document.getElementById('editModal').style.display = 'block';
        }
        
        function closeModal() {
            document.getElementById('editModal').style.display = 'none';
        }
        
        function saveEdit() {
            var form = document.getElementById('editForm');
            var inputs = form.querySelectorAll('input');
            var setClause = [];
            var whereClause = '';
            var pkValue = '';
            
            // 判断是否为数字的函数
            function isNumeric(value) {
                if (value === '' || value === null || value === undefined) return false;
                return !isNaN(value) && isFinite(value) && value.trim() !== '';
            }
            
            for (var i = 0; i < inputs.length; i++) {
                var name = inputs[i].name;
                var value = inputs[i].value.trim();
                var isPk = inputs[i].disabled;
                
                if (isPk) {
                    pkValue = value;
                    // 检查主键是否为数字
                    if (isNumeric(value)) {
                        whereClause = '[' + name + '] = ' + value;
                    } else {
                        whereClause = '[' + name + '] = \'' + value.replace(/'/g, "''") + '\'';
                    }
                } else {
                    // 检查字段值是否为数字
                    if (isNumeric(value)) {
                        setClause.push('[' + name + '] = ' + value);
                    } else {
                        setClause.push('[' + name + '] = \'' + value.replace(/'/g, "''") + '\'');
                    }
                }
            }
            
            if (setClause.length === 0) {
                alert('没有可编辑的字段');
                return;
            }
            
            var sql = 'UPDATE [' + currentEditTable + '] SET ' + setClause.join(', ') + ' WHERE ' + whereClause;
            
            console.log('执行 SQL:', sql);
            
            if (confirm('确定要执行以下 SQL 更新吗？\n\n' + sql)) {
                // 发送到服务器执行
                var form = document.createElement('form');
                form.method = 'post';
                form.action = '?panel=structure&table=' + encodeURIComponent(currentEditTable);
                form.style.display = 'none';
                
                var input = document.createElement('input');
                input.name = 'sqlquery';
                input.value = sql;
                form.appendChild(input);
                
                document.body.appendChild(form);
                form.submit();
            }
        }
        
        // 点击模态框外部关闭
        document.getElementById('editModal').addEventListener('click', function(e) {
            if (e.target === this) {
                closeModal();
            }
        });
        
        // Base64 编码函数（移除填充字符）
        function encodeBase64() {
            var text = document.getElementById('sqlText').value;
            var encoded = btoa(unescape(encodeURIComponent(text)));
            // 移除末尾的填充字符 =
            encoded = encoded.replace(/=+$/, '');
            document.getElementById('sqlBase64').value = encoded;
        }
        
        // Base64 解码函数
        function decodeBase64() {
            var encoded = document.getElementById('sqlBase64').value;
            try {
                var decoded = decodeURIComponent(escape(atob(encoded)));
                document.getElementById('sqlText').value = decoded;
            } catch(e) {
                alert('Base64 解码失败，请检查输入');
            }
        }
        
        // 复制 Base64 到剪贴板
        function copyBase64() {
            var base64 = document.getElementById('sqlBase64').value;
            navigator.clipboard.writeText(base64).then(function() {
                alert('Base64 已复制到剪贴板');
            }).catch(function() {
                alert('复制失败，请手动复制');
            });
        }
        
        // 自动更新 Base64（移除填充字符）
        function updateBase64() {
            var text = document.getElementById('sqlText').value;
            if (text) {
                var encoded = btoa(unescape(encodeURIComponent(text)));
                // 移除末尾的填充字符 =
                encoded = encoded.replace(/=+$/, '');
                document.getElementById('sqlBase64').value = encoded;
            }
        }
        
        // 页面加载时初始化
        document.addEventListener('DOMContentLoaded', function() {
            updateBase64();
        });
        
        // 提交查询前确保 Base64 已生成
        function submitQuery() {
            var sqlText = document.getElementById('sqlText').value;
            if (!sqlText.trim()) {
                alert('请输入 SQL 语句');
                return false;
            }
            updateBase64();
            var base64 = document.getElementById('sqlBase64').value;
            if (!base64.trim()) {
                alert('Base64 编码失败');
                return false;
            }
            return true;
        }
    </script>
    
    <%
    ' 清理连接
    If isConnected Then
        conn.Close
        Set conn = Nothing
    End If
    %>
    <% End If ' 认证检查结束 %>
</body>
</html>
<%
Function IIf(condition, truePart, falsePart)
    If condition Then
        IIf = truePart
    Else
        IIf = falsePart
    End If
End Function

' Base64 解码函数（自动补齐填充字符）
Function Base64Decode(ByVal base64String)
    ' 自动补齐填充字符 =
    Dim padding
    padding = 4 - (Len(base64String) Mod 4)
    If padding <> 4 Then
        base64String = base64String & String(padding, "=")
    End If
    
    Dim objXML, objNode
    Set objXML = Server.CreateObject("MSXML2.DOMDocument")
    Set objNode = objXML.createElement("b64")
    objNode.DataType = "bin.base64"
    objNode.Text = base64String
    Base64Decode = StreamToString(objNode.nodeTypedValue)
    Set objNode = Nothing
    Set objXML = Nothing
End Function

' 字节流转字符串
Function StreamToString(ByVal bytes)
    Dim objStream
    Set objStream = Server.CreateObject("ADODB.Stream")
    objStream.Type = 1 ' adTypeBinary
    objStream.Open
    objStream.Write bytes
    objStream.Position = 0
    objStream.Type = 2 ' adTypeText
    objStream.Charset = "utf-8"
    StreamToString = objStream.ReadText
    objStream.Close
    Set objStream = Nothing
End Function
%>
