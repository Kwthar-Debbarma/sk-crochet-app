from flask import Flask, render_template, request
import openpyxl

app = Flask(__name__)

@app.route('/register', methods=['GET', 'POST'])
def register():
    if request.method == 'POST':
        name = request.form['name']
        email = request.form['email']
        password = request.form['password']
        confirm_password = request.form['confirmPassword']

        # --- In a real application, you'd validate and hash the password here ---
        # For now, we are just storing them directly.
        # You should add validation (e.g., password matching, email format) here.

        try:
            # Check if the Excel file exists, if not, create it with headers
            try:
                workbook = openpyxl.load_workbook('users.xlsx')
                sheet = workbook.active
            except FileNotFoundError:
                workbook = openpyxl.Workbook()
                sheet = workbook.active
                # Add headers to the new sheet
                sheet.append(['Name', 'Email', 'Password'])

            # Append the user data as a new row
            sheet.append([name, email, password])

            # Save the Excel file
            workbook.save('users.xlsx')

            # Return a success message to the user
            return "Registration successful! User data saved to Excel."
        except Exception as e:
            # Return an error message if something goes wrong during file operations
            return f"An error occurred: {e}"
    else:
        # For GET requests, render the registration form
        return render_template('register.html')

if __name__ == '__main__':
    # This part is for local development only and won't run on PythonAnywhere
    app.run(debug=True)