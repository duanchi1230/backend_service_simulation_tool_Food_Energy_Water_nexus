from model import app
import win32com.client
WEAP = win32com.client.Dispatch("WEAP.WEAPApplication")
value = WEAP.ResultValue(
			"\Supply and Resources\Transmission Links\\to Municipal\\from CAPWithdral:Flow[m^3]", 1986, 1, "5% Population Growth", 1986, 12, "Average")

print(value)
if __name__ == "__main__":
    app.run(debug=True)
    print(__name__)