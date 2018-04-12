<!--#include file="PostmenException.asp"-->
<!--#include file="../../../../../resources/inc/make_json.asp"-->

<%
' local vars
Dim api_key, region, config, myPostmen

api_key = "1234"
region = "sandbox"
set config = Server.CreateObject("Scripting.Dictionary")
config.add "endpoint", "blah-blah-blah"

set myPostmen = (new Postmen)(api_key, region, config)

'**
'* Class Handler
'*
'* @package Postmen
'**
Class Postmen
  ' local vars
  Private cv_isConstructed
  Private cv_api_key, cv_version, cv_error, cv_config

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
    cv_config.add "safe", false
    cv_config.add "proxy", Server.CreateObject("Scripting.Dictionary")
    set cv_config = MergeDicts(config)
    ' set attributes concerning rate limiting and auto-retry
    cv_delay = 1
    cv_retries = 0
    cv_max_retries = 5
    cv_calls_left = null
  End Function

  Public Function buildXmlHttpParams(method, path, config)
    isContructed()
    ' local vars
    dim parameters, url, query, xmlhttp_params, headers

    set parameters = MergeDicts(config)

    if isNull(parameters("body")) then
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

    query = null
    if not isNull(parameters("query")) then
      query = parameters("query")
    end if

    url = generateURL(parameters("endpoint"), path, method, query)

    set xmlhttp_params = server.createobject("scripting.dictionary")
    xmlhttp_params.add "url", url
    xmlhttp_params.add "customrequest", method
    xmlhttp_params.add "httpheaders", headers

    if not method = "GET" then
      xmlhttp_params.add "postfields" = parameters("body")
    end if

    set buildXmlHttpParams = xmlhttp_params
  End Function

  Public Function myCall(method, path, config)
    isContructed()
    'body
  End Function

  Public Function processCurlResponse(fv_response, parameters)
    isContructed()
    'body
  End Function

  Public Function handleError(err_message, err_code, err_retryable, err_details, parameters)
    Dim error

    ' NEED TO ADD CODE HERE, IF WE END UP USING THIS

    handleError = null
  End Function

  Public Function handle(parsed, parameters)
    isContructed()
    'body
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

    config("query") = query

    callGET = myCall("GET", path, config)
  End Function

  Public Function callPOST(path, body, config)
    isContructed()

    config("body") = body

    callPOST = myCall("POST", path, config)
  End Function

  Public Function callPUT(path, body, config)
    isContructed()

    config("body") = body
    callPUT = myCall("PUT", path, config)
  End Function

  Public Function callDELETE(path, body, config)
    isContructed()

    config("body") = body

    callDELETE = myCall("DELETE", path, config)
  End Function

  Public Function myGet(resource, id, query, config)
    isContructed()

    if not id = null then
      myGet = callGET("/v3/"&resource&"/"&id, query, config)
    else
      myGet = callGET("/v3/"&resource, query, config)
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

    create = callPOST("/v3/"&resource, payload, config)
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
