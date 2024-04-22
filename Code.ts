
const users = listAllUsers()

const sessionId = conscriboLogin();

function createUsers() {
  const people = getPeople();
  for (const person of people) {
    let user = createNewUser(person.voornaam, person.achternaam, person.email, person.code);
    if (user) {
      try {
        AdminDirectory.Users.insert(user);
        console.log(`User ${person.voornaam} ${person.achternaam} created in G Suite`);
      } catch (e) {
        console.error(`Error creating user ${person.voornaam} ${person.achternaam}: ${e}`);
      }
    } else {
      console.log(`User ${person.voornaam} ${person.achternaam} already exists in G Suite`);
    }

  }
}

function conscriboLogin(): string {
  let loginRequestData = makeConscriboRequest("authenticateWithUserAndPass", {
    userName: PropertiesService.getScriptProperties().getProperty("conscribo.username"),
    passPhrase: PropertiesService.getScriptProperties().getProperty("conscribo.password"),
  })
  let res = UrlFetchApp.fetch(PropertiesService.getScriptProperties().getProperty("conscribo.url"), {
    method: "post",
    payload: JSON.stringify(loginRequestData),
    contentType: "application/json",
    headers: {
      "X-Conscribo-API-Version": "0.20161212"
    }
  });
  if (res.getResponseCode() === 200) {
    let data = JSON.parse(res.getContentText());
    if (data.result.success) {
      return data.result.sessionId;
    } else {
      for (const notification of data.result.notifications) {
        console.error(notification);
      }
    }
  }
}

function listUsers() {
  for (const user of users) {
    console.log(JSON.stringify(user))
  }
}

function listFieldDefinitions() {
  let listFieldDefinitions = makeConscriboRequest("listFieldDefinitions", {
    entityType: "lid"
  });
  let res = UrlFetchApp.fetch(PropertiesService.getScriptProperties().getProperty("conscribo.url"), {
    method: "post",
    payload: JSON.stringify(listFieldDefinitions),
    contentType: "application/json",
    headers: {
      "X-Conscribo-API-Version": "0.20161212",
      "X-Conscribo-SessionId": sessionId
    }
  });
  if (res.getResponseCode() === 200) {
    let data = JSON.parse(res.getContentText());
    if (data.result.success) {
      for (const field of data.result.fields) {
        console.log(field);
      }
    } else {
      for (const notification of data.result.notifications) {
        console.error(notification);
      }
    }
  }
}

interface Person {
  voornaam: string,
  achternaam: string,
  email: string,
  code: string,
  membership_ended: string

}

function getPeople(): Person[] {
  let listRelations = makeConscriboRequest("listRelations", {
    entityType: "lid",
    requestedFields: {
      fieldName: ["voornaam", "achternaam", "email", "code", "membership_ended"]
    }
  });
  let res = UrlFetchApp.fetch(PropertiesService.getScriptProperties().getProperty("conscribo.url"), {
    method: "post",
    payload: JSON.stringify(listRelations),
    contentType: "application/json",
    headers: {
      "X-Conscribo-API-Version": "0.20161212",
      "X-Conscribo-SessionId": sessionId
    }
  });
  if (res.getResponseCode() === 200) {
    let data = JSON.parse(res.getContentText());
    if (data.result.success) {
      console.log(data.result.resultCount);
      let people = data.result.relations;
      let actualPeople = [];
      for (const key in people) {
        if (people.hasOwnProperty(key)) {
          const person = people[key];
          actualPeople.push(person);
        }
      }
      return actualPeople;
    } else {
      for (const notification of data.result.notifications) {
        console.error(notification);
      }
    }
  }
}

function makeConscriboRequest(command: string, data: any) {
  return {
    request: {
      command,
      ...data
    }
  }
}

function listAllUsers(): GoogleAppsScript.AdminDirectory.Schema.User[] {
  let pageToken: any;
  let page: GoogleAppsScript.AdminDirectory.Schema.Users;
  let allUsers = [];
  do {
    page = AdminDirectory.Users.list({
      domain: 'aegee-delft.nl',
      orderBy: 'givenName',
      maxResults: 100,
      pageToken: pageToken
    });
    const users = page.users;
    if (!users) {
      console.log('No users found.');
      return allUsers;
    }

    allUsers = allUsers.concat(users)

    pageToken = page.nextPageToken;
  } while (pageToken);
  return allUsers;
}

function createNewUser(firstName: string, lastName: string, backupEmail: string, id: string): GoogleAppsScript.AdminDirectory.Schema.User | undefined {
  if (users.find(user => user.name.givenName === firstName && user.name.familyName === lastName)) {
    console.log(`User ${firstName} ${lastName} already exists in G Suite`);
    return;
  }
  let primaryEmail = `${firstName.replace(/\s/g, "")}${lastName.replace(/\s/g, "")}@aegee-delft.nl`.toLocaleLowerCase("nl-NL");
  console.log(`Creating user ${firstName} ${lastName} with email ${primaryEmail}`);
  let user = {
    primaryEmail,
    name: {
      givenName: firstName,
      familyName: lastName,
    },
    password: Math.random().toString(36),
    recoveryEmail: backupEmail,
    changePasswordAtNextLogin: true,
    externalIds: [{
      type: "custom",
      customType: "conscriboID",
      value: id
    }],
    orgUnitPath: "/Member",
  }
  return user;
}
