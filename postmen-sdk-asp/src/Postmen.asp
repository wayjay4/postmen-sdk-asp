<!--#include file="PostmenException.asp"-->

<%
api_key = "1234"
region = "sandbox"
set config = Server.CreateObject("Scripting.Dictionary")
config.add "endpoint", "blah-blah-blah"

set batman = (new Postmen)(api_key, region, config)

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

  Public default Function construct(api_key, region, config)
    set construct = me
    cv_isConstructed = true

    ' set all the context attributes
    if api_key = "" then
      err.raise 60000, "PostmenException" ,"API key is required"
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
    set cv_config = MergeDicts(cv_config, config)
    ' set attributes concerning rate limiting and auto-retry
    cv_delay = 1
    cv_retries = 0
    cv_max_retries = 5
    cv_calls_left = null
  End Function




  ' HELPER FUNCTIONS
  Private Function isContructed()
    if (not cv_isConstructed) then
      err.raise 60000, "ObjectNotConstructedException", "Postmen is not constructed"
    end if
  End Function

  Private Function MergeDicts(dct1, dct2)
    ' Merge 2 dictionaries. The second dictionary will override the first if they have the same key
    ' local vars
    Dim result, allKeys1, allKeys2, i, key

    ' initialize result object
    Set result = CreateObject("Scripting.Dictionary")

    ' put dct1 object in result
    allKeys1 = dct1.Keys  ' get all the keys into an array
    For i = 0 To dct1.Count - 1   ' iterate through the object
      myKey = allKeys1(i) ' this is the key value
      result.add myKey, dct1(myKey) ' add key/value pair to result
    Next

    ' put dct2 object in result
    allKeys2 = dct2.Keys   ' get all the keys into an array
    For i = 0 To dct2.Count - 1 ' iterate through the second object
      myKey = allKeys2(i)   ' this is the key value

      ' add key/value pair to result
      If result.Exists(myKey) Then
        result.item(myKey) = dct2(myKey)
      Else
        result.add myKey, dct2(myKey)
      End If
    Next

    ' return result
    Set MergeDicts = result
  End Function

  Private Sub Class_Terminate(  )
    'On Nothingd
  End Sub
End Class


%>
