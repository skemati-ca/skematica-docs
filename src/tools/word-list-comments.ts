import { validateDocxPath } from '../validation.js';

export const WORD_LIST_COMMENTS_SCHEMA = {
  type: 'object',
  properties: {
    filePath: { type: 'string', description: 'Absolute path to the .docx file' },
    section: { type: 'string', description: 'Limit to comments within a specific section' },
    status: { type: 'string', enum: ['all', 'resolved', 'unresolved'], description: 'Filter by resolution status. Default: "all"' },
  },
  required: ['filePath'],
} as const;

export async function wordListComments(args: Record<string, unknown>): Promise<Record<string, unknown>> {
  const { filePath, status = 'all' } = args as { filePath: string; section?: string; status?: string };

  const err = validateDocxPath(filePath);
  if (err) return { content: [{ type: 'text', text: err }], isError: true };

  const { DocxDocument } = await import('../docx.js');
  const doc = await DocxDocument.load(filePath);
  let comments = await doc.getComments();

  // Filter by status
  if (status === 'resolved') {
    comments = comments.filter((c: { isResolved: boolean }) => c.isResolved);
  } else if (status === 'unresolved') {
    comments = comments.filter((c: { isResolved: boolean }) => !c.isResolved);
  }

  return {
    content: [{
      type: 'text',
      text: JSON.stringify({
        totalComments: countAll(comments),
        comments,
        _suggestions: {
          word_create_comment: { tool: 'word_create_comment', description: 'Add a new comment' },
          word_reply_to_comment: { tool: 'word_reply_to_comment', description: 'Reply to a comment' },
          word_resolve_comment: { tool: 'word_resolve_comment', description: 'Mark a comment as resolved' },
        },
      }, null, 2),
    }],
  };
}

function countAll(comments: Array<{ replies: unknown[] }>): number {
  let count = 0;
  for (const c of comments) {
    count += 1 + c.replies.length;
  }
  return count;
}
