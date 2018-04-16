<!--#include file="PostmenException.asp"-->
<!--#include file="../../../../../resources/inc/make_json.asp"-->
<!--#include file="../../../../../resources/inc/parse_json.asp"-->

<%
' local vars
Dim api_key, region, config, api_myPostmen, result, key

api_key = "1234"
region = "sandbox"
set config = Server.CreateObject("Scripting.Dictionary")

set api_myPostmen = (new Postmen)(api_key, region, config)

resource = "labels"
id = null
set query = Server.CreateObject("Scripting.Dictionary")
set config = Server.CreateObject("Scripting.Dictionary")
set result = api_myPostmen.myGet(resource, id, query, config)

if api_myPostmen.Error() then
  set errorMessage = api_myPostmen.ErrorMessage
  errorMessage = "error code: "&errorMessage("err_code")&"<br />error message: "&errorMessage("err_message")&""
  response.write "<p>"&errorMessage&"</p>"
  response.write "<p>Response: "&result("strResponse")&"</p>"
elseif result.Exists("strResponse") then
  response.write result("strResponse")
elseif result.Exists("data") then
  response.write "<p>"
  response.write "next_token: "& result("data")("next_token") & "<br>"
  response.write "limit: "& result("data")("limit") & "<br>"
  response.write "created_at_min: "& result("data")("created_at_min") & "<br>"
  response.write "created_at_max: "& result("data")("created_at_max") & "<br>"
  'response.write "labels: "& result("data")("labels") & "<br>"
  response.write "</p>"
end if


'**
'* Class Handler
'*
'* @package Postmen
'**
Class Postmen
  ' local vars
  Private cv_isConstructed
  Private cv_api_key, cv_version, cv_config, cv_error, cv_error_details

  ' auto-retry if retryable attributes
  Private cv_retry, cv_delay, cv_retries, cv_max_retries, cv_calls_left

  ' rate limiting attributes
  Private cv_rate

  Private Sub Class_Initialize( )
    'Constructor
    cv_isConstructed = false
    cv_api_key = null
    cv_version = null
    cv_error = null
    cv_config = null
    cv_retry = null
    cv_delay = null
    cv_retries = null
    cv_max_retries = null
    cv_calls_left = null
  End Sub

  Public Default Function construct(api_key, region, config)
    set construct = me
    cv_isConstructed = true

    ' set all the context attributes
    if api_key = "" then
      err.raise 60001, "PostmenException" ,"API key is required"
    end if

    cv_version = "1.0.0"
    cv_api_key = api_key
    set cv_config = Server.CreateObject("Scripting.Dictionary")
    cv_config.add "endpoint", "https://"&region&"-api.postmen.com"
    cv_config.add "retry", true
    cv_config.add "rate", true
    cv_config.add "array", false
    cv_config.add "raw", false
    cv_config.add "safe", true
    set cv_config = MergeDicts(config)
    ' set attributes concerning rate limiting and auto-retry
    cv_delay = 1
    cv_retries = 0
    cv_max_retries = 5
    cv_calls_left = null
    cv_error = false
    set cv_error_details = CreateObject("Scripting.Dictionary")
  End Function

  Public Property Get Error()
    Error = cv_error
  End Property

  Public Property Get ErrorMessage()
    set ErrorMessage = cv_error_details
  End Property

  Public Function buildXmlHttpParams(method, path, config)
    isContructed()
    ' local vars
    dim parameters, url, query, xmlhttp_params, headers

    set parameters = MergeDicts(config)

    if isNull(parameters("body")) OR isEmpty(parameters("body")) then
      parameters("body") = ""
    elseif not TypeName(parameters("body")) = "String" then
      if (parameters("body").Count - 1) = 0 then
        parameters("body") = ""
      else
        parameters("body") = (new JSON)(empty, parameters("body"), true)
      end if
    end if

    set headers = server.createobject("scripting.dictionary")
    headers.add "Content-Type", "application/json"
    headers.add "postmen-api-key", cv_api_key
    headers.add "x-postmen-agent", "php-sdk-"&ScriptEngineMinorVersion

    set query = Server.CreateObject("Scripting.Dictionary")
    if not (parameters("query").Count - 1) = 0 then
      set query = parameters("query")
    end if

    url = generateURL(parameters("endpoint"), path, method, query)

    set xmlhttp_params = server.createobject("scripting.dictionary")
    xmlhttp_params.add "url", url
    xmlhttp_params.add "customrequest", method
    xmlhttp_params.add "httpheaders", headers

    if not method = "GET" then
      xmlhttp_params.add "postfields", parameters("body")
    else
      xmlhttp_params.add "postfields", null
    end if

    set buildXmlHttpParams = xmlhttp_params
  End Function

  Public Function myCall(method, path, config)
    isContructed()
    ' local vars
    dim xmlhttp, parameters, retry, raw, safe, xmlhttp_params, strStatus, strResponse

    cv_retries = cv_retries + 1

    set parameters = MergeDicts(config)

    if isNull(method) then
      method = parameters("method")
    else
      parameters("method") = method
    end if

    if isNull(path) then
      path = parameters("path")
    else
      parameters("path") = path
    end if

    retry = parameters("retry")
    raw = parameters("raw")
    safe = parameters("safe")

    set xmlhttp_params = buildXmlHttpParams(method, path, parameters)

    set xmlhttp = server.createobject("Microsoft.XMLHTTP")
    xmlhttp.open xmlhttp_params("customrequest"), xmlhttp_params("url"), false

    for each key in xmlhttp_params("httpheaders")
      header = xmlhttp_params("httpheaders")(key)
      'Response.Write objPayment.Name
      xmlhttp.setRequestHeader key, header

      'response.write "key: "&key&", value: "& xmlhttp_params("httpheaders")(key) &VbCrLf
    next

    if not isNull(xmlhttp_params("postfields")) then
      xmlhttp.send xmlhttp_params("postfields")
    else
      xmlhttp.send
    end if

    strStatus = xmlhttp.Status
    strResponse = xmlhttp.ResponseText

    set myCall = processXmlHttpResponse(strStatus, strResponse, parameters)
  End Function

  Public Function processXmlHttpResponse(strStatus, strResponse, parameters)
    isContructed()
    ' local vars
    dim json_parcer, parsed, raw_response, err_message, err_code, err_retryable, err_details

    ' instantiate the class
    set json_parcer = New JSONobject
    ' parce the string object strResponse
    set parsed = json_parcer.Parse(strResponse)

    if not isObject(parsed) then
      if parameters("raw") then
        set raw_response = server.CreateObject("Scripting.Dictionary")
        raw_response.add "strResponse", strResponse
        set processXmlHttpResponse = raw_response
      else
        set processXmlHttpResponse = handle(parsed, parameters)
      end if
    else
      err_message = "Something went wrong on Postmen's end."
      err_code = 500
      err_retryable = false
      set err_details = CreateObject("Scripting.Dictionary")

      set processXmlHttpResponse = handleError(err_message, err_code, err_retryable, err_details, parameters)
    end if
  End Function

  Public Function handleError(err_message, err_code, err_retryable, err_details, parameters)
    Dim result, errorMessage

    cv_error = true
    cv_error_details.add "err_message", err_message
    cv_error_details.add "err_code", err_code
    cv_error_details.add "err_details", err_details

    set result = CreateObject("Scripting.Dictionary")
    result.add "strResponse", null

    if not parameters("safe") then
      set handleError = result
    else
      errorMessage = "error code: "&err_code&", error message: "&err_message
      err.raise 60001, "PostmentErrorException", errorMessage
      set handleError = result
    end if
  End Function

  Public Function handle(parsed, parameters)
    isContructed()
    ' local vars

    if parsed.value("meta")("code") = 200 then
      ' output the json object
      'parsed.Write()
      ' output a single value from the json object
      'response.write "<p>"
      'response.write "code: "& parsed.value("meta")("code") & "<br>"
      'response.write "message: "& parsed.value("meta")("message") & "<br>"
      'response.write "details: "& parsed.value("meta")("details") & "<br>"
      'response.write "next_token: "& parsed.value("data")("next_token") & "<br>"
      'response.write "limit: "& parsed.value("data")("limit") & "<br>"
      'response.write "created_at_min: "& parsed.value("data")("created_at_min") & "<br>"
      'response.write "created_at_max: "& parsed.value("data")("created_at_max") & "<br>"
      'response.write "labels: "& parsed.value("data")("labels") & "<br>"
      'response.write "</p>"

      set result = CreateObject("Scripting.Dictionary")
      result.add "data", parsed.value("data")

      set handle = result
    else
      'NEEDS ERROR CATCHING CODE HERE
      response.write "Postmen server side error occurred."
    end if
  End Function

  ' allow query as a string
  Public Function generateURL(url, path, method, query)
    isContructed()

    if method = "GET" then
      if TypeName(query) = "String" then
        if Len(query) > 0 then
          Dim firstLetter : firstLetter = Left(query, 1)
          if firstLetter = "?" then
            generateURL = url & path & query
          else
            generateURL = url & path & "?" & query
          end if
        end if
      end if

      if not isNull(query) then
        Dim qr : qr = http_build_query(query)
        if Len(qr) > 0 then
          generateURL = url & path & "?" & qr
        end if
      end if
    end if

    generateURL = url & path
  End Function

  Public Function handleRetry(parameters)
    isContructed()

    if cv_retries < cv_max_retries then
      mySleep(cv_delay)

      cv_delay = cv_delay * 2

      handleRetry = myCall(null, null, parameters)
    else
      cv_retries = 0
      cv_delay = 1

      handleRetry = null
    end if
  End Function

  Public Function callGET(path, query, config)
    isContructed()

    config.add "query", query

    set callGET = myCall("GET", path, config)
  End Function

  Public Function callPOST(path, body, config)
    isContructed()

    config("body") = body

    set callPOST = myCall("POST", path, config)
  End Function

  Public Function callPUT(path, body, config)
    isContructed()

    config("body") = body
    set callPUT = myCall("PUT", path, config)
  End Function

  Public Function callDELETE(path, body, config)
    isContructed()

    config("body") = body

    set callDELETE = myCall("DELETE", path, config)
  End Function

  Public Function myGet(resource, id, query, config)
    isContructed()

    if not id = null then
      set myGet = callGET("/v3/"&resource&"/"&id, query, config)
    else
      set myGet = callGET("/v3/"&resource, query, config)
    end if
  End Function

  Public Function create(resource, payload, config)
    isContructed()

    if not TypeName(payload) = "String" then
      'set payload = Server.CreateObject("Scripting.Dictionary")
      'payload.add "async", false

      myKey = "async"
      if payload.Exists(myKey) then
        payload.item(myKey) = false
      else
        payload.add myKey, false
      end if
    end if

    set create = callPOST("/v3/"&resource, payload, config)
  End Function

  Public Function getError()
    isContructed()

    ' return error
    getError = cv_error
  End Function



  ' HELPER FUNCTIONS
  Private Function isContructed()
    if (not cv_isConstructed) then
      err.raise 60000, "ObjectNotConstructedException", "Postmen is not constructed"
    end if
  End Function

  '** takes a dictonary object config as parameter
  '*  returns merged dictionary object
  '*  values from config are priority
  '**
  Private Function MergeDicts(config)
    ' Merge 2 dictionaries. The second dictionary will override the first if they have the same key (i.e: values from config are prioritary)
    ' local vars
    Dim result, allKeys1, allKeys2, i, key

    ' initialize result object
    Set result = CreateObject("Scripting.Dictionary")

    ' put cv_config object in result
    allKeys1 = cv_config.Keys  ' get all the keys into an array
    For i = 0 To cv_config.Count - 1   ' iterate through the object
      myKey = allKeys1(i) ' this is the key value
      result.add myKey, cv_config(myKey) ' add key/value pair to result
    Next

    ' put config object in result
    allKeys2 = config.Keys   ' get all the keys into an array
    For i = 0 To config.Count - 1 ' iterate through the second object
      myKey = allKeys2(i)   ' this is the key value

      ' add key/value pair to result
      If result.Exists(myKey) Then
        result.item(myKey) = config(myKey)
      Else
        result.add myKey, config(myKey)
      End If
    Next

    ' return result
    Set MergeDicts = result
  End Function

  Private Function mySleep(seconds)
    Dim DateTimeResume
    DateTimeResume= DateAdd("s", NumberOfSeconds, Now())
    Do Until (Now() > DateTimeResume)
      ' do nothing, but wait
    Loop
  End Function

  Private Function http_build_query(queryObj)
    ' make code here to parse a dictionary object to a html query
    Dim resultQuery, allKeys1, i, myKey

    ' initialize resultQuery
    resultQuery = ""

    ' make queryObj into string var resultQuery
    allKeys1 = queryObj.Keys  ' get all the keys into an array
    For i = 0 To queryObj.Count - 1   ' iterate through the object
      myKey = allKeys1(i) ' this is the key value

      if resultQuery = "" then
        resultQuery = myKey&"="&Server.URLEncode(queryObj(myKey))
      else
        resultQuery = resultQuery & "&" & myKey&"="&Server.URLEncode(queryObj(myKey))
      end if
    Next

    http_build_query = resultQuery
  End Function

  Private Sub Class_Terminate(  )
    'On Nothingd
  End Sub
End Class


%>
