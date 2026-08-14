sub init()
    m.api = "http://10.1.10.190:5050/"
end sub

function getProfileGridConfig()
    url = m.api + "profileGridConfig"
    getContent(url)
end function

function getProfileSliderData()
    url = m.api + "profileSliderData"
    getContent(url)
end function

function getMenuGridConfig()
    url = m.api + "menuGridConfig"
    getContent(url)
end function

function getRowListData()
    url = m.api + "rowItemConfigWithSizes"
    getContent(url)
end function

function getContent(link)
    urlObject = CreateObject("roUrlTransfer")
    urlObject.setRequest("GET")
    urlObject.SetUrl(link)
    port = CreateObject("roMessagePort")
    urlObject.SetMessagePort(port)

    if urlObject.AsyncGetToString()
        msg = wait(0, port)
        if type(msg) = "roUrlEvent"
            responseData = msg.getString()
            ' print responseData
            response = ParseJson(responseData)
            m.top.responseData = response
        end if
    end if
    
end function