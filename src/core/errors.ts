export type ErrorCode =
  | 'UNAVAILABLE'
  | 'EMPTY_SELECTION'
  | 'STALE_SELECTION'
  | 'INVALID_PLAN'
  | 'BUSY'
  | 'EXPIRED_TASK'
  | 'HOST_CONNECTION'
  | 'SELECTION_TOO_LARGE';

export class HelperError extends Error {
  constructor(
    public readonly code: ErrorCode,
    message: string,
  ) {
    super(message);
    this.name = 'HelperError';
  }
}

export function userMessage(error: unknown): string {
  if (error instanceof HelperError) return error.message;
  return '这一步没有完成。请检查文件是否打开，再试一次。';
}
