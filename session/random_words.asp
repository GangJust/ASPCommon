<%
	'一个随机字生成方法
	function getRandomWords(length)
		dim wordsStr
		dim wordsLength

		wordsStr = ""
		wordsLength = (26*2+10)

		dim words(62) '大小写字母，加上数字

		'受其他编程语言影响，我就将数组长度当做 length-1 以免记混

		'大写字母A-Z
		for i = 0 to (wordsLength - (26+10))
			words(i) = chr(65+i)
		next

		'小写字母a-z
		for i = 26 to (wordsLength - 10)
			words(i) = chr(97+(i-26))
		next

		'数字
		for i = 52 to wordsLength-1
			words(i) = chr(48+(i-52))
		next

		
		'生成随机字
		for i = 0 to length-1

			dim rand,randWord

			randomize timer
			rand = int((rnd() * wordsLength))

			randWord = words(rand)

			wordsStr = wordsStr & randWord
		next

		getRandomWords = wordsStr
	end function
%>