const RECEIVING_EMAIL = "<email that receives the trustpilot reviews>"
const SLACK_WEBHOOK_URL = "<incomming webhook url from slack>"

const slackTemplate = (title, content, link, stars) => ({
  "blocks": [
    {
      "type": "context",
      "elements": [
        {
          "type": "mrkdwn",
          "text": range(stars, () => ":star:").join('') + " " + title
        }
      ]
    },
    {
      "type": "context",
      "elements": [
        {
          "type": "plain_text",
          "text": content,
        },
      ]
    },
    {
      "type": "context",
      "elements": [

        {
          type: `mrkdwn`,
          text: `<${link}|See this review and reply>`,
        },
      ]
    }
  ]
})


function triggerNewTrustPilotEmails() {
  var threads = GmailApp.search(`from:Trustpilot list:(${RECEIVING_EMAIL})`)
  for (const thread of threads) {
    const messages = thread.getMessages()

    for (const message of messages) {
      if (message.isStarred()) {
        var text = message.getPlainBody().split("\n")

        // get review title
        var titleIndex = text.findIndex((v) => v.indexOf('just left a new') != -1)
        var title = text[titleIndex]

        // get review content
        var line = titleIndex + 1
        var content = []
        while (text[line].indexOf('See this review and reply') == -1) {
          content.push(text[line])
          line++
        }

        // get review link & stars
        var linkIndex = line + 1
        var link = text[linkIndex].replace('[', '').replace(']', '').trim()
        var stars = Number(title.match(/\d-star/gm)[0].split('-star')[0])
        
        // build & send message to slack
        var template = slackTemplate(title, content.join('').trim(), link, stars)
        UrlFetchApp.fetch(SLACK_WEBHOOK_URL, {
          method: 'post',
          payload: JSON.stringify(template)
        })
        message.unstar()
      }
    }
  }
}

function range(n, fn) {
  var a = []
  for (var i = 0; i < n; i++) {
    a.push(fn())
  }

  return a
}
