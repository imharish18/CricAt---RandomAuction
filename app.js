async function readExcelFile(i,value) {
      const response = await fetch('data/respose.xlsx');
      const data = await response.arrayBuffer();
      const workbook = XLSX.read(data, { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const sheetData = XLSX.utils.sheet_to_json(sheet);


      if(value=="name"){
        return sheetData[i].Name;
      }
      else if(value=="role"){
         return sheetData[i].Role;
      }
      else if(value=="style"){
        return sheetData[i].Style;
      }
      else if(value=="link"){
        return sheetData[i].link;
      }

}
function firstLast(name) {
    const parts = name.trim().split(" ");
    
    const capitalize = (word) => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();

    const firstName = parts.length > 0 ? capitalize(parts[0]) : "";
    const lastName = parts.length > 1 ? capitalize(parts[parts.length - 1]) : "";
  
    return { firstName, lastName };
  }
// WORKING ID
function extractFileId(driveLink2) {

  const parts = driveLink2.split('id=');
  if (parts.length > 1) {
    return parts[1];
  } else {
    return null;
  }
}
let defaultLink = "https://drive.google.com/thumbnail?id="
function generateThumbnailUrl(driveLink) {
  console.log("Original"+driveLink);
  console.log("Code Only"+extractFileId(driveLink))
  console.log("Defalt"+defaultLink);
  console.log(defaultLink+extractFileId(driveLink))
  return defaultLink + extractFileId(driveLink);

}
let firstName = document.querySelector("#firstName");
let lastName = document.querySelector("#lastName");
let role = document.querySelector("#batBowl");
let style = document.querySelector("#style");
let link = "I am link"
let playerCount = -1;
let limg = document.querySelector("#navLimg");
let rimg = document.querySelector("#navRimg");
limg.addEventListener("click",()=>{
  if(teamCount.length == 6){
    removeLock();
  }
    if(playerCount==0){
        playerCount=0;
        doAll(playerCount);
    }
    else{
        playerCount-=1;
        doAll(playerCount);
    }
})
rimg.addEventListener("click",()=>{
  if(teamCount.length == 6){
    removeLock();
  }
    playerCount+=1;
    console.log(playerCount);
    doAll(playerCount);
})

function doAll(playerCount){
    readExcelFile(playerCount, "name").then((result) => {
        firstName.textContent = firstLast(result).firstName;
        lastName.textContent = firstLast(result).lastName;
      }).catch((error) => {
        console.error('Error:', error);
    });
    readExcelFile(playerCount, "role").then((result) => {
        role.textContent = result;
    
      }).catch((error) => {
        console.error('Error:', error);
    });
    readExcelFile(playerCount, "style").then((result) => {
        style.textContent = result;
      }).catch((error) => {
        console.error('Error:', error);
    });

    (async () => {
        try {
          link = await readExcelFile(playerCount, "link");
          link = generateThumbnailUrl(link);
          let linkIn = document.querySelector("#imageSync")
          linkIn.src = link;
        } catch (error) {
          console.error("Error:", error);
        }
    })();
}
//  RANDOM SELECTION
let start = document.querySelector("#imgStart");
let team1 = document.querySelector("#team1");
let team2 = document.querySelector("#team2");
let team3 = document.querySelector("#team3");
let team4 = document.querySelector("#team4");
let team5 = document.querySelector("#team5");
let team6 = document.querySelector("#team6");
let teamCount = [];
function removeLock(){
  team1.classList.remove('done');
  team2.classList.remove('done');
  team3.classList.remove('done');
  team4.classList.remove('done');
  team5.classList.remove('done');
  team6.classList.remove('done');
}
function selectRandomTeam() {
  if (teamCount.length === 6) {
    teamCount = [];
  }

  let randomNumber;
  do {
    randomNumber = Math.floor(Math.random() * 6); 
  } while (teamCount.includes(randomNumber));
  
  teamCount.push(randomNumber);
  return randomNumber;
}
function animateTeams(callback) {
  let world = 0;
  const teams = [team1, team2, team3, team4, team5, team6];
  const speed = [120, 120, 110, 100, 90, 80, 70, 60, 50, 50, 50, 50, 60, 70, 80, 90, 100, 110, 120];
  let teamIndex = 0;
  function animateNextTeam() {
    if (world >= 6 ) {
      if (typeof callback === "function") callback();
      return;
    }

    while (teams[teamIndex].classList.contains('done')) {
      teamIndex++;
      if (teamIndex >= teams.length) {
        teamIndex = 0;
        world++;
      }
      if (world >= 6) {
        if (typeof callback === "function") callback();
        return;
      }
    }
    teams[teamIndex].classList.add('excited');

    setTimeout(() => {
      teams[teamIndex].classList.remove('excited');
      teamIndex++;
      if (teamIndex >= teams.length) {
        teamIndex = 0;
        world++;
      }
      
      animateNextTeam();
    }, speed[Math.min(world, speed.length - 1)]);
  }
  animateNextTeam();
}
function freezeTeam(team){
  if (team == 0) {
    team1.classList.add('excited');
    team1.classList.add('done');

    setTimeout(function() {
      team1.classList.remove('excited');
      teamPosterAnimation("csk");
    }, 2000);
  }
  if (team == 1) {
    team2.classList.add('excited');
    team2.classList.add('done');

    setTimeout(function() {
      team2.classList.remove('excited');
      teamPosterAnimation("rcb");
    }, 2000);
  }
  if (team == 2) {
    team3.classList.add('excited');
    team3.classList.add('done');

    setTimeout(function() {
      team3.classList.remove('excited');
      teamPosterAnimation("kkr");
    }, 2000);
  }
  if (team == 3) {
    team4.classList.add('excited');
    team4.classList.add('done');

    setTimeout(function() {
      team4.classList.remove('excited');
      teamPosterAnimation("mi");
    }, 2000);
  }
  if (team == 4) {
    team5.classList.add('excited');
    team5.classList.add('done');

    setTimeout(function() {
      team5.classList.remove('excited');
      teamPosterAnimation("rr");
    }, 2000);
  }
  if (team == 5) {
    team6.classList.add('excited');
    team6.classList.add('done');

    setTimeout(function() {
      team6.classList.remove('excited');
      teamPosterAnimation("srh");
    }, 2000);
  }
}
let tapSound = new Audio("sound/srhAnthem.mp3");
start.addEventListener("click", () => {
  tapSound = new Audio("sound/srhAnthem.mp3");
  tapSound.pause();
  tapSound.currentTime = 0; 
  tapSound.play().catch(error => console.error("Audio playback error:", error));
  if (teamCount.length == 6) {
    removeLock();
  }
  animateTeams(() => {
    freezeTeam(selectRandomTeam());
  });
});
function teamPosterAnimation(name){
  tapSound.pause();
  tapSound = new Audio("sound/celebration.mp3");
  tapSound.play().catch(error => console.error("Audio playback error:", error));
  work1();
  work2();
  work3();
  work4();
  let popUpImg = document.querySelector("#popUpImage");
  document.getElementById("popUpImg").src = `teamPoster/${name}.png`;
  popUpImg.style.display = "block"
  setTimeout(() => {
  popUpImg.style.display = "none"
  }, 7000);
}
function work4(){


  var pumpkin = confetti.shapeFromPath({
    path: 'M449.4 142c-5 0-10 .3-15 1a183 183 0 0 0-66.9-19.1V87.5a17.5 17.5 0 1 0-35 0v36.4a183 183 0 0 0-67 19c-4.9-.6-9.9-1-14.8-1C170.3 142 105 219.6 105 315s65.3 173 145.7 173c5 0 10-.3 14.8-1a184.7 184.7 0 0 0 169 0c4.9.7 9.9 1 14.9 1 80.3 0 145.6-77.6 145.6-173s-65.3-173-145.7-173zm-220 138 27.4-40.4a11.6 11.6 0 0 1 16.4-2.7l54.7 40.3a11.3 11.3 0 0 1-7 20.3H239a11.3 11.3 0 0 1-9.6-17.5zM444 383.8l-43.7 17.5a17.7 17.7 0 0 1-13 0l-37.3-15-37.2 15a17.8 17.8 0 0 1-13 0L256 383.8a17.5 17.5 0 0 1 13-32.6l37.3 15 37.2-15c4.2-1.6 8.8-1.6 13 0l37.3 15 37.2-15a17.5 17.5 0 0 1 13 32.6zm17-86.3h-82a11.3 11.3 0 0 1-6.9-20.4l54.7-40.3a11.6 11.6 0 0 1 16.4 2.8l27.4 40.4a11.3 11.3 0 0 1-9.6 17.5z',
    matrix: [0.020491803278688523, 0, 0, 0.020491803278688523, -7.172131147540983, -5.9016393442622945]
  });
  // tree shape from https://thenounproject.com/icon/pine-tree-1471679/
  var tree = confetti.shapeFromPath({
    path: 'M120 240c-41,14 -91,18 -120,1 29,-10 57,-22 81,-40 -18,2 -37,3 -55,-3 25,-14 48,-30 66,-51 -11,5 -26,8 -45,7 20,-14 40,-30 57,-49 -13,1 -26,2 -38,-1 18,-11 35,-25 51,-43 -13,3 -24,5 -35,6 21,-19 40,-41 53,-67 14,26 32,48 54,67 -11,-1 -23,-3 -35,-6 15,18 32,32 51,43 -13,3 -26,2 -38,1 17,19 36,35 56,49 -19,1 -33,-2 -45,-7 19,21 42,37 67,51 -19,6 -37,5 -56,3 25,18 53,30 82,40 -30,17 -79,13 -120,-1l0 41 -31 0 0 -41z',
    matrix: [0.03597122302158273, 0, 0, 0.03597122302158273, -4.856115107913669, -5.071942446043165]
  });
  // heart shape from https://thenounproject.com/icon/heart-1545381/
  var heart = confetti.shapeFromPath({
    path: 'M167 72c19,-38 37,-56 75,-56 42,0 76,33 76,75 0,76 -76,151 -151,227 -76,-76 -151,-151 -151,-227 0,-42 33,-75 75,-75 38,0 57,18 76,56z',
    matrix: [0.03333333333333333, 0, 0, 0.03333333333333333, -5.566666666666666, -5.533333333333333]
  });
  
  var defaults = {
    scalar: 2,
    spread: 180,
    particleCount: 30,
    origin: { y: -0.1 },
    startVelocity: -35,
  };
  
  confetti({
    ...defaults,
    shapes: [pumpkin],
    colors: ['#ff9a00', '#ff7400', '#ff4d00']

  });
  confetti({
    ...defaults,
    shapes: [tree],
    colors: ['#8d960f', '#be0f10', '#445404']
  });
  confetti({
    ...defaults,
    shapes: [heart],
    colors: ['#f93963', '#a10864', '#ee0b93']
  });
}
function work3(){
  var end = Date.now() + (3000);

// go Buckeyes!
var colors = ["#cd252e", '#193A8A'];

(function frame() {
  confetti({
    particleCount: 5,
    angle: 60,
    spread: 60,
    origin: { x: 0 },
    colors: colors
  });
  confetti({
    particleCount: 5,
    angle: 120,
    spread: 60,
    origin: { x: 1 },
    colors: colors
  });

  if (Date.now() < end) {
    requestAnimationFrame(frame);
  }
}());
}
function work2(){
  var duration = 3000;
var animationEnd = Date.now() + duration;
var skew = 1;

function randomInRange(min, max) {
  return Math.random() * (max - min) + min;
}

(function frame() {
  var timeLeft = animationEnd - Date.now();
  var ticks = Math.max(200, 200 * (timeLeft / duration));
  skew = Math.max(0.8, skew - 0.001);

  confetti({
    particleCount: 1,
    startVelocity: 0,
    ticks: ticks,
    origin: {
      x: Math.random(),
      // since particles fall down, skew start toward the top
      y: (Math.random() * skew) - 0.2
    },
    colors: ['#E31A81'],
    shapes: ['heart'],
    gravity: randomInRange(0.4, 0.6),
    scalar: randomInRange(0.4, 1),
    drift: randomInRange(-0.4, 0.4)
  });

  if (timeLeft > 0) {
    requestAnimationFrame(frame);
  }
}());
}
function work1(){
  var duration = 3000;
  var animationEnd = Date.now() + duration;
  var defaults = { startVelocity: 30, spread: 360, ticks: 40, zIndex: 2000 };

function randomInRange(min, max) {
  return Math.random() * (max - min) + min;
}

var interval = setInterval(function() {
  var timeLeft = animationEnd - Date.now();

  if (timeLeft <= 0) {
    return clearInterval(interval);
  }

  var particleCount = 700 * (timeLeft / duration);
  // since particles fall down, start a bit higher than random
  confetti({ ...defaults, particleCount, origin: { x: randomInRange(0.1, 0.3), y: Math.random() - 0.2 } });
  confetti({ ...defaults, particleCount, origin: { x: randomInRange(0.7, 0.9), y: Math.random() - 0.2 } });
}, 500);
}
