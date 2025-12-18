import sys
import os

# Aseguramos que Python encuentre el módulo sgos_web
sys.path.append(os.path.abspath(os.path.dirname(__file__)))

from sgos_web.app import app, db, User

def main():
    while True:
        print("\n=== GESTIÓN DE USUARIOS SGOS ===")
        print("1. Crear nuevo usuario")
        print("2. Listar usuarios")
        print("3. Eliminar usuario")
        print("4. Salir")
        
        opcion = input("\nSelecciona una opción (1-4): ")
        
        with app.app_context():
            if opcion == "1":
                username = input("Nuevo Usuario: ").strip()
                if not username:
                    print("El usuario no puede estar vacío.")
                    continue
                    
                if User.query.filter_by(username=username).first():
                    print(f"¡Error! El usuario '{username}' ya existe.")
                    continue
                    
                password = input("Contraseña: ").strip()
                if not password:
                    print("La contraseña no puede estar vacía.")
                    continue
                    
                user = User(username=username)
                user.set_password(password)
                db.session.add(user)
                db.session.commit()
                print(f"✅ Usuario '{username}' creado exitosamente.")
                
            elif opcion == "2":
                users = User.query.all()
                print("\n👥 Usuarios registrados:")
                for u in users:
                    print(f"- {u.username} (ID: {u.id})")
                    
            elif opcion == "3":
                username = input("Usuario a eliminar: ").strip()
                if username == "admin":
                    print("⚠️ No puedes eliminar al admin principal por seguridad.")
                    continue
                    
                user = User.query.filter_by(username=username).first()
                if user:
                    db.session.delete(user)
                    db.session.commit()
                    print(f"🗑️ Usuario '{username}' eliminado.")
                else:
                    print("❌ Usuario no encontrado.")
            
            elif opcion == "4":
                print("¡Hasta luego!")
                break
            
            else:
                print("Opción no válida.")

if __name__ == "__main__":
    main()
