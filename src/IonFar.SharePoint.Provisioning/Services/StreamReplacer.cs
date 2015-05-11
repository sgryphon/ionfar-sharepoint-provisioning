using System.IO;

namespace IonFar.SharePoint.Provisioning.Services
{
    public class StreamReplacer
    {
        public string ReplaceValue(string filePath, string oldValue, string newValue)
        {
            var reader = new StreamReader(filePath);
            var fileAsString = reader.ReadToEnd().Replace(oldValue, newValue);
            reader.Close();

            var streamWriter = new StreamWriter(filePath, false);
            streamWriter.Write(fileAsString);
            streamWriter.Flush();
            streamWriter.Close();

            return fileAsString;
        }
    }
}
