const dotenv = require("dotenv");
const request = require("request");
const fs = require("fs");
var mime = require("mime");
dotenv.config();

let file_to_upload;
let clientFolderName;
let ABSOLUTE_PATH;

let url =
  process.env.AUTHORITY_URL + process.env.TENANT_ID + "/oauth2/v2.0/token";

// With client credentials flows permissions need to be granted in the portal by a tenant administrator.
// The scope is always in the format "<resource>/.default"
exports.uploadFilesToOneDrive = async function (args) {
  //var mainLogic = async function () {
  return new Promise(async (resolve, reject) => {
    file_to_upload = args.filename;
    clientFolderName = args.folderName;
    ABSOLUTE_PATH = args.path;
    try {
      let resolveResponse = await getToken(url);
      //authorize(JSON.parse(content), uploadFileToFolder);
      resolve(resolveResponse);
    } catch (err) {
      return reject(err);
    }
  });
};

/**
 * Acquires token with client credentials.
 * @param {object} clientRequest
 */
// Get Authentication Token
async function getToken(clientRequest) {
  return new Promise(async (resolve, reject) => {
    try {
      let options = {
        method: "POST",
        url: clientRequest,
        headers: {
          "Postman-Token": "d33cfc11-3ee8-4f0e-980d-8a4629b048ac",
          "cache-control": "no-cache",
          "Content-Type": "application/x-www-form-urlencoded",
        },
        form: {
          client_id: process.env.CLIENT_ID,
          scope: process.env.GRAPH_ENDPOINT + ".default",
          client_secret: process.env.CLIENT_SECRET,
          grant_type: process.env.GRANT_TYPE,
        },
      };

      request(options, async function (error, response, body) {
        if (error) throw new Error(error);
        let res = JSON.parse(body);
        let token = res.access_token;

        let resonseUpload = await uploadFileToDriveFolder(token);
        resolve(resonseUpload);
      });
    } catch (err) {
      reject(err);
    }
  });
}

// Check Whether Folder Exist in Drive
function checkAndCreateFolder(token) {
  return new Promise(async (resolve, reject) => {
    let folderDetails = await getAllFolders(token, clientFolderName);
    // console.log("FOlder Details:", folderDetails);
    if (folderDetails != undefined || folderDetails != null) {
      resolve(folderDetails.id);
    } else {
      try {
        let createdFolderId = await createFolder(token, clientFolderName);
        resolve(createdFolderId.id);
      } catch (err) {
        console.log("Error in createFolder", err.toString());
        reject(err.toString());
      }
    }
  });
}

// Create Folder In Drive
function createFolder(token, folderName) {
  return new Promise((resolve, reject) => {
    let url =
      process.env.GRAPH_ENDPOINT +
      "v1.0/drives/" +
      process.env.DRIVE_ID +
      "/items/" +
      process.env.PARENT_FOLDER_ID +
      "/children";
    let options = {
      method: "POST",
      url: url,
      headers: {
        Authorization: "Bearer " + token,
      },
      body: {
        name: folderName,
        folder: {},
      },
      json: true,
    };

    request(options, function (error, response, body) {
      if (error) throw new Error(error);

      resolve(body);
    });
  });
}

// Upload File To Drive
function uploadFileToDriveFolder(token) {
  return new Promise(async (resolve, reject) => {
    let subFolderId;

    console.log("ABSOLUTE_PATH", ABSOLUTE_PATH);

    let fileName = file_to_upload;

    //Check if folder for client exists
    try {
      subFolderId = await checkAndCreateFolder(token);
    } catch (err) {
      reject(err);
    }

    var options = {
      method: "PUT",
      url:
        process.env.GRAPH_ENDPOINT +
        "/v1.0/drives/" +
        process.env.DRIVE_ID +
        "/items/" +
        subFolderId +
        ":/" +
        fileName +
        ":/content",
      headers: {
        "Postman-Token": "04ecc268-49a6-4f7e-b0a6-7dceb00ba067",
        "cache-control": "no-cache",
        "Content-Type": mime.getType(fileName),
        Authorization: "Bearer " + token,
      },
      body: fs.createReadStream(ABSOLUTE_PATH),
    };

    request(options, function (error, response, body) {
      if (error) {
        // Handle error
        console.log("Error in upload to drive", JSON.stringify(error));
        reject(error);
      } else {
        console.log("Uploaded to drive succefully");
        resolve(true);
      }
    });
  });
}

// List All Folders From OneDrive
async function getAllFolders(token, clientFolder) {
  return new Promise((resolve, reject) => {
    let url =
      process.env.GRAPH_ENDPOINT +
      "v1.0/drives/" +
      process.env.DRIVE_ID +
      "/items/" +
      process.env.PARENT_FOLDER_ID +
      "/children";
    let options = {
      method: "GET",
      url: url,
      headers: {
        Authorization: "Bearer " + token,
      },
    };
    request(options, function (error, response, body) {
      if (error) throw new Error(error);
      let folderList = JSON.parse(body);
      let folderNames = [];
      if (folderList.value.length > 0) {
        for (let i = 0; i < folderList.value.length; i++) {
          folderNames.push(folderList.value[i].name);
        }
        if (folderNames.includes(clientFolder)) {
          let val = folderNames.indexOf(clientFolder);
          resolve(folderList.value[val]);
        } else {
          resolve(null);
        }
      } else {
        resolve(null);
      }
    });
  });
}
