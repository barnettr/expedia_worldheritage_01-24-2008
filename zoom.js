
var stateArray = [];
stateArray['United States of America'] = 3;
stateArray['Canada'] = 3;
stateArray['Australia'] = 3;
stateArray['France'] = 5;
stateArray['Italy'] = 5;
stateArray['Sweden'] = 5;
stateArray['Turkey'] = 5;
stateArray['Ukraine'] = 5;
stateArray['Russian Federation'] = 2;
stateArray['United Kingdom of Great Britain and Northern Ireland'] = 5;
stateArray['Argentina'] = 3;
stateArray['Brazil'] = 3;
stateArray['Chile'] = 4;
stateArray['Ecuador'] = 4;
stateArray['Mexico'] = 4;
stateArray['China'] = 3;
stateArray['Iran (Islamic Republic of)'] = 5;
stateArray['Japan'] = 5;
stateArray['Pakistan'] = 5;
stateArray['India'] = 4;
stateArray['Algeria'] = 5;
stateArray['Libyan Arab Jamahiriya'] = 5;
stateArray['Germany'] = 5;
stateArray['Poland'] = 5;
stateArray['Portugal'] = 4;
stateArray['Colombia'] = 5;
stateArray['Venezuela (Bolivarian Republic of)'] = 5;
stateArray['Indonesia'] = 5;
stateArray['New Zealand'] = 5;
stateArray['Philippines'] = 5;
stateArray['South Africa'] = 5;

function displayZoomLevel(stateArray,state) {
  if (typeof stateArray[state] == 'undefined') {
    zoom = 6;
    }else {
       zoom = stateArray[state];
  }
  return zoom;
}

