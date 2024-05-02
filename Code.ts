
type MyUser = GoogleAppsScript.AdminDirectory.Schema.User & { recoveryEmail: string };

const users = listAllUsers()

const sessionId = conscriboLogin();

function createUsers() {
  const people = getPeople();
  for (const person of people) {
    const existingUser = findUser(person.code);
    if (existingUser) {
      console.log(`User ${existingUser.name.fullName} already exists in G Suite with code ${person.code}`);
      updateUser(existingUser, person);
    } else {
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
}

function findUser(code: string): MyUser | undefined {
  return users.find(user => user.externalIds && user.externalIds.find((id: GoogleAppsScript.AdminDirectory.Schema.UserExternalId) => id.customType === "conscriboID" && id.value === code));
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
  // for (const user of users.filter(user => user.externalIds == null)) {
  //   console.log(user.name.fullName);
  // }
  let people = getPeople();
  let me = people.find(person => person.code === "1016");
  console.log(JSON.stringify(me));
  let u2 = users.find(user => user.primaryEmail === "juliusdejeu@aegee-delft.nl");
  console.log(JSON.stringify(u2));
  updateUser(u2, me);
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

function listAllUsers(): MyUser[] {
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

function userConscriboId(user: GoogleAppsScript.AdminDirectory.Schema.User): string | undefined {
  let externalIds: GoogleAppsScript.AdminDirectory.Schema.UserExternalId[] = user.externalIds;
  if (externalIds) {
    let id = externalIds.find((id: GoogleAppsScript.AdminDirectory.Schema.UserExternalId) => id.customType === "conscriboID");
    if (id) {
      return id.value;
    }
  }
}

function updateUser(user: MyUser, cUser: Person) {
  let reMail = user.recoveryEmail ?? "";
  if (cUser.email === reMail && userConscriboId(user) === cUser.code) {
    console.log(`User ${user.name.fullName} already has the correct email address and conscribo ID`);
    return;
  }
  try {
    let u2 = {
      name: {
        givenName: cUser.voornaam,
        familyName: cUser.achternaam,
      },
      recoveryEmail: cUser.email,
      externalIds: [{
        type: "custom",
        customType: "conscriboID",
        value: cUser.code
      }],
    }
    AdminDirectory.Users.update(u2, user.primaryEmail)
  } catch (e) {
    console.error(`Error updating user ${user.name.fullName}: ${e}`);
  }
}
