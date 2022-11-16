// const msal = require("@azure/msal-node");
const dotenv = require("dotenv");
const request = require("request");
const fs = require("fs");
var mime = require("mime");
var async = require('async');
// const { uploadFilesToOneDrive } = require("./uploadFileToOneDrive");
dotenv.config();

let file_to_upload;
let clientFolderName;
let ABSOLUTE_PATH;

let url =
  process.env.AUTHORITY_URL + process.env.TENANT_ID + "/oauth2/v2.0/token";

/**
 * Configuration  Object need to be passed to MSAL instance on creation
 */
//  const config = {
//   headers: {
//     "Content-Type": "application/x-www-form-urlencoded",
//   },
//     form: {
//         client_id: process.env.CLIENT_ID,
//         client_secret: process.env.CLIENT_SECRET
//         grant_type:process.env.GRANT_TYPE,
//         scope:process.env.GRAPH_ENDPOINT+".default"
//    }
// };

// Create msal application object

// With client credentials flows permissions need to be granted in the portal by a tenant administrator.
// The scope is always in the format "<resource>/.default"
async function uploadFilesToOneDrive(args) {
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
  // return await cca.acquireTokenByClientCredential(clientRequest);
}

function checkAndCreateFolder(token) {
  return new Promise(async (resolve, reject) => {
    let folderDetails = await getAllFolders(token, clientFolderName);
    // console.log("FOlder Details:", folderDetails);
    if (folderDetails != undefined || folderDetails != null) {
      resolve(folderDetails.id);
    } else {
      try {
        let createdFolderId = await createFolder(token, clientFolderName);
        // console.log("createdFolder is", createdFolderId);
        resolve(createdFolderId.id);
      } catch (err) {
        console.log("Error in createFolder", err.toString());
        reject(err.toString());
      }
    }
  });
}

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
    // console.log(options);

    request(options, function (error, response, body) {
      if (error) throw new Error(error);

      // console.log(body)
      resolve(body);
    });
  });
}

function uploadFileToDriveFolder(token) {
  return new Promise(async (resolve, reject) => {
    let subFolderId;

    console.log("ABSOLUTE_PATH", ABSOLUTE_PATH);

    // Let's change the file name to IST from UTC
    let fileName = file_to_upload;

    //Check if folder for client exists
    try {
      subFolderId = await checkAndCreateFolder(token);
    } catch (err) {
      reject(err);
    }
    console.log("Sub FolderId:", subFolderId);
    //   fs.readFile(absolute_path, function read(e, f) {
    //     request.put({
    //         url: 'https://graph.microsoft.com/v1.0/drives/'+process.env.DRIVE_ID+"/items/"+process.env.PARENT_FOLDER_ID+"/"+clientFolderName + ':/children/' + fileName + ':/content',
    //         headers: {
    //             'Authorization': "Bearer " + token,
    //             'Content-Type': mime.getType(absolute_path), // When you use old version, please modify this to "mime.lookup(file)",
    //         },
    //         body: f,
    //     }, function(er, re, bo) {
    //         console.log(bo);
    //     });
    // });

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
        ":/createUploadSession",
      headers: {
        "Postman-Token": "04ecc268-49a6-4f7e-b0a6-7dceb00ba067",
        "cache-control": "no-cache",
        "Content-Type": "application/json",
        Authorization: "Bearer " + token,
      },
      body: '{"item": {"@microsoft.graph.conflictBehavior": "rename", "name": "' + fileName + '"}}',
    };

    request(options, function (error, response, body) {
      if (error) {
        // Handle error
        console.log("Error in upload to drive", JSON.stringify(error));
        reject(error);
      } else {
        uploadFile(JSON.parse(body).uploadUrl)
        console.log(
          body,
          " Uploaded to drive succefully-------------------------------------------------------------------------------------------------------------------",
        );
        resolve(true);
      }
    });
  });
}

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
    // console.log(clientFolder)
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

function uploadFile(uploadUrl) { // Here, it uploads the file by every chunk.
    async.eachSeries(getparams(), function(st, callback){
        setTimeout(function() {
            fs.readFile(args.filename, function read(e, f) {
                request.put({
                    url: uploadUrl,
                    headers: {
                        'Content-Length': st.clen,
                        'Content-Range': st.cr,
                    },
                    body: f.slice(st.bstart, st.bend + 1),
                }, function(er, re, bo) {
                    console.log(bo);
                });
            });
            callback();
        }, st.stime);
    });
}

function getparams(){
    var allsize = fs.statSync(args.filename).size;
    var sep = allsize < (60 * 1024 * 1024) ? allsize : (60 * 1024 * 1024) - 1;
    var ar = [];
    for (var i = 0; i < allsize; i += sep) {
        var bstart = i;
        var bend = i + sep - 1 < allsize ? i + sep - 1 : allsize - 1;
        var cr = 'bytes ' + bstart + '-' + bend + '/' + allsize;
        var clen = bend != allsize - 1 ? sep : allsize - i;
        var stime = allsize < (60 * 1024 * 1024) ? 5000 : 10000;
        ar.push({
            bstart : bstart,
            bend : bend,
            cr : cr,
            clen : clen,
            stime: stime,
        });
    }
    return ar;
}

