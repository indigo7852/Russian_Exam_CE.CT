export function shuffle(items) {
  const result=[...items];
  for(let i=result.length-1;i>0;i--){const j=Math.floor(Math.random()*(i+1));[result[i],result[j]]=[result[j],result[i]];}
  return result;
}
export function scoreTask(task, selected) {
  const correct=task.words.filter(w=>w.correct==="А").map(w=>w.position);
  const mistakes=new Set([...correct.filter(x=>!selected.includes(x)),...selected.filter(x=>!correct.includes(x))]).size;
  return {correct, mistakes, points:mistakes===0?2:mistakes===1?1:0};
}
export const TEST_SCORE_TABLE=[0,1,2,3,4,5,6,7,9,10,12,14,16,18,20,22,23,25,27,28,29,31,32,34,35,36,37,38,39,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67,68,69,70,71,72,73,74,75,76,77,78,79,80,81,82,83,84,85,87,89,90,92,96,98,100];
export function primaryToTest(primary, maxPrimary) {
  // Пропорциональный эквивалент для короткой тренировки; официально шкала применяется к полной работе.
  const scaled=Math.max(0,Math.min(80,Math.round(primary/maxPrimary*80)));
  return {scaledPrimary:scaled,testScore:TEST_SCORE_TABLE[scaled]};
}
export function formatTime(seconds){return `${String(Math.floor(seconds/60)).padStart(2,"0")}:${String(seconds%60).padStart(2,"0")}`;}
