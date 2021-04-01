import { data } from "./data/output";

/**
 * Update with your own base URL and replace the rest of the URL to match your
 * own up to the point where the custom folder names come in. So here I have
 * 'Programs' and 'ABCXYZ' to replace in two places, and this assumes that the
 * name of my SharePoint document library is 'Project Documents'.
 */
const baseURL = `https://YOUR_BASE_URL_HERE/sites/Programs/ABCXYZ/Project%20Documents/Forms/AllItems.aspx?RootFolder=%2Fsites%2FPrograms%2FABCXYZ%2FProject%20Documents%2F`;
let errors: string[] = [];

const splitInputIntoArray = (inputString: string) => {
  return inputString.split("\n");
};

const removeWColonSlash = (inputString: string[]): string[] => {
  return inputString.map((line) => line.substring(3));
};

const removeBlanksAndForms = (inputString: string[]) => {
  return inputString.filter((line) => {
    if (line === "") return false;
    if (line.substr(0, 5) === "Forms") return false;
    return true;
  });
};

const splitPathIntoSections = (inputString: string[]): string[][] => {
  return inputString.map((line) => {
    return line.split("\\");
  });
};

interface TaggedObjectWithPROACIDType {
  inputStringArray: string[];
  sortString?: string;
  type: "project" | "area" | "category" | "id" | "unknown";
}

const tagArrayWithPROACIDType = (
  input: string[][]
): TaggedObjectWithPROACIDType[] => {
  // This is a very naïve function. If the input is bad, this returns bad.
  return input.map((item) => {
    switch (item.length) {
      case 1:
        return { inputStringArray: item, type: "project" };
      case 2:
        return { inputStringArray: item, type: "area" };
      case 3:
        return { inputStringArray: item, type: "category" };
      case 4:
        return { inputStringArray: item, type: "id" };
      default:
        return { inputStringArray: item, type: "unknown" };
    }
  });
};

const createSortString = (
  inputArray: TaggedObjectWithPROACIDType[]
): TaggedObjectWithPROACIDType[] => {
  return inputArray.map((item, outerIndex) => {
    // Add the sortString property so we can concat it later.
    const returnObject = { ...item, sortString: "" };

    /** itemReturn.inputStringArray is an array starting with the project, then
     *  area, then category, etc. So it doesn't make any difference what `type`
     *  this thing is, that's not what sortString depends on.
     */
    returnObject.inputStringArray.forEach((pathComponent, index) => {
      switch (index) {
        case 0:
          // This should be the Project -- let's check
          if (/\d\d\d /.test(pathComponent.substr(0, 4))) {
            returnObject.sortString = returnObject.inputStringArray[
              index
            ].substr(0, 3);
          } else {
            errors.push(
              `Line ${outerIndex} thinks it's a project, but it isn't: ${pathComponent}.`
            );
          }
          break;
        case 1:
          // This should be the Area
          if (/\d\d-\d\d /.test(pathComponent.substr(0, 6))) {
            returnObject.sortString += returnObject.inputStringArray[
              index
            ].substr(0, 1);
          } else {
            errors.push(
              `Line ${outerIndex} thinks it's an area, but it isn't: ${pathComponent}.`
            );
          }
          break;
        case 2:
          // This should be the Category
          if (/\d\d /.test(pathComponent.substr(0, 3))) {
            returnObject.sortString += returnObject.inputStringArray[
              index
            ].substr(1, 1);
          } else {
            errors.push(
              `Line ${outerIndex} thinks it's a category, but it isn't: ${pathComponent}.`
            );
          }
          break;
        case 3:
          // This should be the ID
          if (/\d\d\.\d\d /.test(pathComponent.substr(0, 6))) {
            returnObject.sortString += returnObject.inputStringArray[
              index
            ].substr(3, 2);
          } else {
            errors.push(
              `Line ${outerIndex} thinks it's an ID, but it isn't: ${pathComponent}.`
            );
          }
          break;
        default:
          // Whatever, do nothing
          break;
      }
    });

    // itemReturn.sortString needs to be 7 chars long otherwise a project
    // is 101 and whatever else is 1011213 and that won't sort properly
    returnObject.sortString = returnObject.sortString.padEnd(7, "0");
    return returnObject;
  });
};

const sortArrayOnSortString = (
  taggedObject: TaggedObjectWithPROACIDType[]
): TaggedObjectWithPROACIDType[] => {
  return taggedObject.sort(
    (a, b) => Number(a.sortString) - Number(b.sortString)
  );
};

const sortedArrayToHtml = (
  taggedObject: TaggedObjectWithPROACIDType[]
): string => {
  const AMPERSAND = "%26";
  const COMMA = "%2C";
  const DASH = "%2D";
  const FORWARD_SLASH = "%2F";
  const SPACE = "%20";

  // We push each '000 Project' to this array as we encounter it, and use it to
  // build the index at the end.
  let projectsForIndex: string[] = [];

  // Define the start and end of the HTML string, which we use later.
  let htmlStartString = `<div style="font-family: Consolas, monospace;">
  <div><p>Last updated: ${Date().toString()}</p></div>`;
  let htmlEndString = "</div>";

  // Define mainHtml, which is where we add all of the HTML as we iterate
  // through the input.
  let mainHtml = "";

  // Define the function which we'll use to add each file path to the string.
  const addFilePathWithForwardSlashes = (inputStringArray: string[]) => {
    mainHtml = ``;
    inputStringArray.forEach((filePath) => {
      // prettier-ignore
      mainHtml +=
        filePath
          .replaceAll("&", AMPERSAND)
          .replaceAll(",", COMMA)
          .replaceAll("-", DASH)
          .replaceAll("/", FORWARD_SLASH)
          .replaceAll(" ", SPACE)
          + FORWARD_SLASH;
    });
    return mainHtml;
  };

  // Run through the input object and do the needful on each line, adding it
  // to mainHtml.
  taggedObject.forEach((lineOfInput: TaggedObjectWithPROACIDType) => {
    switch (lineOfInput.type) {
      case "project":
        mainHtml += `<hr /><div id="${lineOfInput.inputStringArray[0]}" style="margin: 10px 0 5px 0; font-weight: bold; text-decoration: underline; font-size: 1.2rem">`;
        mainHtml += `<a href="`;
        mainHtml += baseURL;
        mainHtml += addFilePathWithForwardSlashes(lineOfInput.inputStringArray);
        mainHtml += `">${lineOfInput.inputStringArray[0]}</a></div>`;
        // Push a project to the array which we'll use to generate the index.
        projectsForIndex.push(lineOfInput.inputStringArray[0]);
        break;
      case "area":
        mainHtml += `<div style="text-indent: 4ch; font-weight: bold; margin: 5px 0 2px 0;">`;
        mainHtml += `<a href="`;
        mainHtml += baseURL;
        mainHtml += addFilePathWithForwardSlashes(lineOfInput.inputStringArray);
        mainHtml += `">${lineOfInput.inputStringArray[1]}</a></div>`;
        break;
      case "category":
        mainHtml += `<div style="text-indent: 7ch; margin: 2px 0">`;
        mainHtml += `<a href="`;
        mainHtml += baseURL;
        mainHtml += addFilePathWithForwardSlashes(lineOfInput.inputStringArray);
        mainHtml += `">${lineOfInput.inputStringArray[2]}</a></div>`;
        break;
      case "id":
        mainHtml += `<div style="text-indent: 10ch; margin: 2px 0">`;
        mainHtml += `<a href="`;
        mainHtml += baseURL;
        mainHtml += addFilePathWithForwardSlashes(lineOfInput.inputStringArray);
        mainHtml += `">${lineOfInput.inputStringArray[3]}</a></div>`;
        break;
      default:
        break;
    }
  });

  // Create the index
  // projectsForIndex.reverse();
  let indexHtml = "<h2>Index - click to jump</h2><ul>";
  projectsForIndex.forEach((project) => {
    indexHtml += `<li><a href="#${project}">${project}</a></li>`;
  });
  indexHtml += "</ul>";

  // Stitch it all together
  // let returnString = htmlStartString + indexHtml + mainHtml + htmlEndString;
  return htmlStartString + indexHtml + mainHtml + htmlEndString;
  // return returnString;
};

export const App = () => {
  // prettier-ignore
  const final = 
  sortedArrayToHtml(
    sortArrayOnSortString(
      createSortString(
        tagArrayWithPROACIDType(
          splitPathIntoSections(
            removeBlanksAndForms(
              removeWColonSlash(
                splitInputIntoArray(data)
              )
            )
          )
        )
      )
    )
  );

  return (
    <div style={{ margin: "20px" }}>
      {errors.length > 0 ? (
        <>
          <h3>Errors:</h3>
          <ul>
            {errors.map((error, index) => (
              <li key={index}>{error}</li>
            ))}
          </ul>
          <hr />
        </>
      ) : (
        <div style={{ margin: "10px 0" }}>
          Input parsed without errors. (That doesn't mean it's perfect, it just
          means that this thing has very little error checking.)
        </div>
      )}
      <button
        onClick={() => navigator.clipboard.writeText(final)}
        style={{
          padding: "15px",
          fontSize: "2em",
          fontFamily: "sans-serif",
          color: "darkmagenta",
        }}
      >
        Copy this to clipboard
      </button>
      <div dangerouslySetInnerHTML={{ __html: final }} />
    </div>
  );
};
