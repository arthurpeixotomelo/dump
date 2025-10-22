import { load } from "jsr:@std/dotenv";
// import { generate } from "jsr:@std/uuid/unstable-v7";

const env = await load();

// id = generate();

// async function createAPIMethod() {

// }

// const test = await fetch(
//   "https://design.penpot.app/api/rpc/command/get-teams",
//   {
//     "method": "GET",
//     "headers": {
//       "Accept": "application/json",
//       "Content-Type": "application/json",
//       "Authorization": `Token ${env.PENPOT_TOKEN}`,
//     },
//   },
// );

// const data: Array<Record<PropertyKey, unknown>> = await test.json();

// type Typographies = Record<string, Typography>;

// type Typography = {
//   id: string;
//   name: string;
//   path: string;
//   fontId: string;
//   fontSize: string;
//   fontStyle: string;
//   fontFamily: string;
//   fontWeight: string;
//   fontVariantId: string;
//   lineHeight: string;
//   letterSpacing: string;
//   textTransform: string;
//   modifiedAt?: string;
// };

// type Colors = Record<string, Color>;

// type Color = {
//   id: string;
//   name: string;
//   path: string;  
//   color: string;
//   opacity: number;
//   modifiedAt?: string;
// };


// const requestInit: Record<string, string | object> = {
//   "method": "POST",
//   "headers": {
//     "Accept": "application/json",
//     "Content-Type": "application/json",
//     "Authorization": `Token ${env.PENPOT_TOKEN}`,
//   },
//   "body": {},
// };

// requestInit.body = JSON.stringify({ id: env.DESIGN_SYSTEM_ID });

// const getFileData = await fetch(
//   "https://design.penpot.app/api/rpc/command/get-file",
//   requestInit
// ).then((res) => res.json());

// const body: Record<string, string | object> = {
//   "id": env.DESIGN_SYSTEM_ID,
//   "sessionId": generate(),
//   "revn": getFileData.revn,
//   "changes": [],
// };

// const updateFile = await fetch(
//   "https://design.penpot.app/api/rpc/command/update-file",
//     requestInit,
// ).then((res) => res.json());

// Deno.writeTextFile("./getFile.json", JSON.stringify(getFileData, null, 2));
