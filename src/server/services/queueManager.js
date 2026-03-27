import PQueue from 'p-queue';

const CONCURRENCY = parseInt(process.env.CONCURRENCY_LIMIT, 10) || 3;
const MAX_QUEUE_DEPTH = parseInt(process.env.MAX_QUEUE_DEPTH, 10) || 10;

const queue = new PQueue({ concurrency: CONCURRENCY });

export function enqueue(fn) {
  if (queue.pending + queue.size >= CONCURRENCY + MAX_QUEUE_DEPTH) {
    return null; // at capacity
  }
  return queue.add(fn);
}

export function getQueueInfo() {
  return {
    running: queue.pending,
    waiting: queue.size,
    concurrency: CONCURRENCY,
    maxDepth: MAX_QUEUE_DEPTH,
    atCapacity: queue.pending + queue.size >= CONCURRENCY + MAX_QUEUE_DEPTH,
  };
}

// Returns 0-based position in queue (0 = currently running)
export function getQueuePosition() {
  return queue.size;
}
