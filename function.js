const functions = new Map();

window.function = function (code, ...params) {
  code = code.value;
  if (code === undefined) return undefined;

  const functionCode = `async (p1,p2,p3,p4,p5,p6,p7)=>{ ${code} }`;
  const fn = eval(functionCode);
  functions.set(code, fn);

  return fn(...params.map(p => p.value));
}
