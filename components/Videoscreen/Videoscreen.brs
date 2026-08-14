sub init()
    m.videoBackground = m.top.findNode("videoBackground")
    m.videoPlayer = m.top.findNode("videoPlayer")
    m.videoTimer = m.top.findNode("videoTimer")
    m.videoTimer.observeField("fire","setFocus")
end sub

function showVideo(msg as object)
    videoContent = CreateObject("RoSGNode","ContentNode")
    videoData = msg.getData()
    videoContent.url = videoData.url
    videoContent.title = videoData.title
    videoContent.streamformat = videoData.streamformat

    m.videoPlayer.content = videoContent
    m.videoPlayer.control = "play"
    m.videoTimer.control ="start"
end function

function setFocus()
    print "Video is being set as focus"
    m.videoPlayer.setFocus(true)
end function