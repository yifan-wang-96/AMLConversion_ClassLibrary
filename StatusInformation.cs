// Author: Yifan Wang

namespace AMLConversion_ClassLibrary
{
    public class StatusInformation : ObservableObject
    {
        private string _text;
        private int _progress;
        private int _maxTasks;
        private int _currentTaskIndex;
        private bool _isWorking;

        public StatusInformation()
        {
            _text = "undefined";
            _progress = 0;
            _isWorking = false;
        }

        public bool IsWorking
        {
            get => _isWorking;
            set
            {
                _isWorking = value;
                OnPropertyChanged("IsWorking");
            }
        }

        public int Progress
        {
            get => _progress;
            set
            {
                _progress = value;

                if(value == 100)
                    IsWorking = false;
                OnPropertyChanged("Progress");
            }
        }

        public string Text {
            get => _text;
            set
            {
                _text = value;
                OnPropertyChanged("Text");
            }
        }

        public void Set_taskCount(int maxTasks, int currentTaskIndex = 0)
        {
            _maxTasks = maxTasks;

            if (currentTaskIndex != 0)
            {
                _currentTaskIndex = currentTaskIndex;
                Progress = (_currentTaskIndex * 100) / _maxTasks;
            }
            else
            {
                _currentTaskIndex = 0;
                Progress = 0;
            }

            IsWorking = true;
        }

        public int Finish_task(string taskName = null)
        {
            if (_currentTaskIndex < _maxTasks)
                _currentTaskIndex++;
            else
                _currentTaskIndex = _maxTasks;

            Progress = (_currentTaskIndex * 100) / _maxTasks;
            if(taskName != null)
                Text = taskName;

            return _currentTaskIndex;
        }
    }
}
