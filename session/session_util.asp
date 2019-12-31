<!-- #include file="这里填写数据库公用连接文件路径,若不要也行，如果改文件将贯穿整个项目，那么建议你还是在这里引用数据库文件，而后在余下文件中只需引入该文件即可" -->
<!-- #include file="random_words.asp" -->
<%
	const COOKIE_NAME = "WWWSESSION"
%>

<%
	'检查是否为登录状态
	'如果 request 携带有 cookie 则通过 getCookie 方法判断该 cookie 是否存在于 CookieMap 中
	'存在则返回 true ,否则返回 false

	function hasCookie()
		dim arr,cookieId,sesId,resault

		cookieId = request.cookies(COOKIE_NAME)
		arr = getCookie(cookieId)
		resault = false

		if IsArray(arr) then
			if (arr(1) <> "" or cookieId <>"") and (arr(1) = cookieId) then
				resault = true
			end if
		end if
		
		hasCookie = resault
	end function


	'接受一个cookie字符串，如果该cookie存在则返回该cookie，否则返回一个空
	function getCookie(cookie)
		dim resault
		'resault = ""

		if cookie <> "" then
			'Application.staticObjects("CookieMap")
			if CookieMap.exists(cookie) then
				resault = CookieMap.item(cookie)
			end if
		end if

		getCookie = resault
	end function

	'创建 session 设置超时为 30 分钟
	function createSession()

		Session.timeout = 30
		Session.contents(COOKIE_NAME) = getRandomWords(16)

		createSession = Session.contents(COOKIE_NAME)
	end function

	'创建 cookies 强制用户只保存一天
	function createCookie(userId)
		dim sesId
		sesId = createSession()
		response.cookies(COOKIE_NAME).expires = dateAdd("d",1,Date())
		response.cookies(COOKIE_NAME) = sesId

		call saveCookie(userId,sesId)

		createCookie = sesId
	end function

	'将Cookie保存至 CookieMap
	function saveCookie(userId,sesId)

		'Application.staticObjects("CookieMap")
		if CookieMap.exists(sesId) = false then
			CookieMap.add sesId,array(userId,sesId)
		end if

		'如果会话集合中存在相同用户，不同的cookie值 则移除它
		dim arr
		for each var in CookieMap
			arr = CookieMap.item(var)
			if userId = arr(0) and sesId <> arr(1)then
				CookieMap.remove(var)
			end if
		next

	end function

	'删除指定含有键的Cookie
	function removeCookie(sesId)

		if CookieMap.exists(sesId) then
			dim arr
			arr = CookieMap.item(sesId)
			CookieMap.remove(sesId)
			session.contents().remove(COOKIE_NAME)
		end if

	end function

	'使 session 立即过期
	function clearSession()
		session.abandon()
	end function
%>