// domain wide service account with user impersonation enabled
const SERVICE_ACCOUNT_PRIVATE_KEY = "< private key of the service account >"
const SERVICE_ACCOUNT_CLIENT_EMAIL = "< client email of the service account >"

/** The location where the user photos is stored, cdn url. */
const CONFIG_PROFILE_PHOTO_URL = "< cdn url where the photos are located >"

/** The google workspace domain */
const CONFIG_GWORKSPACE_DOMAIN = "< domain  of your google workspace >"

/** The slack webhook to notify of signature updates */
const CONFIG_SLACK_WEBHOOK_URL = "< incomming webhook url of your slack webhook >"

/** The default phone number for the user signature */
const CONFIG_DEFAULT_COMPANY_PHONE_NUMBER = "< your company phone number >"


type UserPhone = { type: string, value: string }
type UserOrganization = { title: string }

function getUserOAuthService(userEmail: string) {
  const service = OAuth2.createService("signatures")
    .setPrivateKey(SERVICE_ACCOUNT_PRIVATE_KEY)
    .setIssuer(SERVICE_ACCOUNT_CLIENT_EMAIL)
    .setSubject(userEmail)
    .setPropertyStore(PropertiesService.getScriptProperties())
    .setParam('access_type', 'offline')
    .setTokenUrl('https://accounts.google.com/o/oauth2/token')
    .setScope('https://www.googleapis.com/auth/gmail.settings.sharing') // access to alias emails
    .setScope('https://www.googleapis.com/auth/gmail.settings.basic') // access to main profile
  
  // force re-authorization
  service.reset()

  return service
}

function execSingleUser() {
  const userEmail = "<email of the user to update>"
  const response = AdminDirectory.Users.list({ domain: CONFIG_GWORKSPACE_DOMAIN })
  for (const user of response.users) {
    if (userEmail == user.primaryEmail) {
      updateUserSignature(user)
    }
  }
}

function execAllUsers() {
  var response = AdminDirectory.Users.list({ domain: CONFIG_GWORKSPACE_DOMAIN })
  for (const user of response.users) {
    updateUserSignature(user)
  }
}


function getUserWorkPhone(user: GoogleAppsScript.AdminDirectory.Schema.User): string | undefined {
  return (user.phones as unknown as UserPhone[])?.find(({ type }) => type == 'work')?.value
}

function getUserJobTitles(user: GoogleAppsScript.AdminDirectory.Schema.User): string | undefined {
  return (user.organizations as unknown as UserOrganization[])?.map(({ title }) => title)?.join(' | ')
}

function updateUserSignature(user: GoogleAppsScript.AdminDirectory.Schema.User) {
  // generate signature template from file
  var t = HtmlService.createTemplateFromFile('template')

  // set template data
  t.name = user.name.fullName

  t.role = getUserJobTitles(user) ?? ""
  t.phone = getUserWorkPhone(user) ?? CONFIG_DEFAULT_COMPANY_PHONE_NUMBER

  // set the profile photo url, make sure your cdn has the file example: "john@examplecorp.com.png" 
  t.photo = `${CONFIG_PROFILE_PHOTO_URL}/${user.primaryEmail}.png`


  // build html
  var html = t.evaluate()

  // get permission to perfom the action as the user
  var service = getUserOAuthService(user.primaryEmail)

  // reset the service to prevent
  service.reset()

  // perform update
  if (service.hasAccess()) {
    var url = `https://www.googleapis.com/gmail/v1/users/me/settings/sendAs/${user.primaryEmail}`
    var payload = { "signature": html.getContent()}
    var data = JSON.stringify(payload)
    var response = UrlFetchApp.fetch(url, {
      method: 'patch',
      payload: data,
      contentType: 'application/json',
      headers: {
        Authorization: `Bearer ${service.getAccessToken()}`
      }
    })
    console.log(response.getContentText())
  }
}
