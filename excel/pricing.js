module.exports = {
  computeStepPrice: (step) => step.quantity ? Math.max(step.min_price || 0, step.price * step.quantity) : 0,
  computeRoomTotal: ({ steps }) => steps.reduce((prev, curr) => prev + computeStepPrice(curr), 0),
  countProjectSteps: ({ version: { rooms } }) => rooms.reduce((prev, curr) => prev + curr.steps.length, 0),
  computeProjectTotal: ({ version: { rooms } }) => rooms.reduce((prev, r) => prev + computeRoomTotal(r), 0),
}