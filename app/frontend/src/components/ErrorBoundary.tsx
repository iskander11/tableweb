import { Component, ReactNode } from 'react';

export class ErrorBoundary extends Component<{ children: ReactNode }, { error: Error | null }> {
  state = { error: null };

  static getDerivedStateFromError(error: Error) {
    return { error };
  }

  render() {
    if (this.state.error) {
      return (
        <div className="min-h-screen flex items-center justify-center bg-gray-50 p-8">
          <div className="bg-white rounded-xl border border-red-200 p-6 max-w-2xl w-full">
            <h2 className="text-red-600 font-semibold text-lg mb-2">Ошибка приложения</h2>
            <pre className="text-xs text-gray-700 bg-gray-100 rounded p-3 overflow-auto whitespace-pre-wrap">
              {(this.state.error as Error).message}
              {'\n\n'}
              {(this.state.error as Error).stack}
            </pre>
          </div>
        </div>
      );
    }
    return this.props.children;
  }
}
